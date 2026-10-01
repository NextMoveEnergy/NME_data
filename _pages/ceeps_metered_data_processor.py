
import json
import math
import time
import zipfile
from collections import defaultdict
from datetime import datetime
from io import BytesIO
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st


API_URL = "https://api.informatika.si/enotna-vstopna-tocka/merilni-podatki/meter-readings"
LOCAL_TZ = ZoneInfo("Europe/Ljubljana")
CHUNK_SIZE = 50
REQUEST_TIMEOUT = 60
MAX_RETRIES = 3
RETRY_CODES = {429, 502, 503, 504}

DISTRIBUTIONS = {
    2: "Celje",
    3: "Ljubljana",
    4: "Maribor",
    6: "Gorenjska",
    7: "Primorska",
}

CATEGORIES = {
    "dobava": "Supply",
    "odkup": "Purchase",
    "obratovalna_podpora": "Support",
}

REFERENCE_COLUMNS = {
    "merilna_tocka",
    "distribucija",
    "naziv_placnika",
}


class AppError(Exception):
    pass


def authorization(ceeps_id: str) -> str:
    key = "encoded_string_sfa" if ceeps_id == "SFA" else "encoded_string_nme"
    try:
        value = str(st.secrets[key]).strip()
    except Exception as exc:
        raise AppError(f"Missing Streamlit secret '{key}'.") from exc
    if not value:
        raise AppError(f"Streamlit secret '{key}' is empty.")
    return f"Basic {value}"


def request_params(
    message_type: str,
    usage_points: list[str],
    start_date,
    end_date,
) -> list[tuple[str, str]]:
    if message_type == "Daily 15 minute":
        params = [("messageType", "D1_15MIN")]
    elif message_type == "Monthly 15 minute":
        params = [("messageType", "M1_15MIN")]
    elif message_type == "Specify date":
        params = [
            ("startTime", str(start_date)),
            ("endTime", str(end_date)),
        ]
    else:
        raise AppError(f"Unsupported message type: {message_type}")

    params.extend(("usagePoints", point) for point in usage_points)
    return params


def retrieve_chunk(
    message_type: str,
    ceeps_id: str,
    usage_points: list[str],
    start_date,
    end_date,
) -> requests.Response:
    headers = {
        "Accept": "application/json",
        "Authorization": authorization(ceeps_id),
    }

    params = request_params(
        message_type,
        usage_points,
        start_date,
        end_date,
    )

    last_response = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(
                API_URL,
                headers=headers,
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            last_response = response
        except requests.Timeout as exc:
            if attempt == MAX_RETRIES:
                raise AppError(
                    f"CEEPS request timed out after {REQUEST_TIMEOUT} seconds."
                ) from exc
            time.sleep(attempt * 2)
            continue
        except requests.ConnectionError as exc:
            if attempt == MAX_RETRIES:
                raise AppError("Could not connect to the CEEPS API.") from exc
            time.sleep(attempt * 2)
            continue
        except requests.RequestException as exc:
            raise AppError(f"CEEPS request failed: {exc}") from exc

        if response.status_code not in RETRY_CODES:
            return response

        if attempt < MAX_RETRIES:
            retry_after = response.headers.get("Retry-After")
            try:
                delay = int(retry_after) if retry_after else attempt * 2
            except ValueError:
                delay = attempt * 2
            time.sleep(min(delay, 10))

    if last_response is None:
        raise AppError("CEEPS did not return a response.")
    return last_response


def http_error(response: requests.Response) -> str:
    messages = {
        400: "Invalid request.",
        401: "Authentication failed.",
        403: "Access denied.",
        404: "Resource not found.",
        429: "API rate limit reached.",
        500: "CEEPS internal server error.",
        502: "CEEPS gateway error.",
        503: "CEEPS service unavailable.",
        504: "CEEPS gateway timeout.",
    }

    body = (response.text or "<empty response>").strip()
    try:
        body = json.dumps(response.json(), ensure_ascii=False, indent=2)
    except (ValueError, requests.exceptions.JSONDecodeError):
        pass

    return (
        f"{messages.get(response.status_code, 'CEEPS request failed')} "
        f"HTTP {response.status_code} {response.reason}. "
        f"Response: {body[:2000]}"
    )


def parse_response(response: requests.Response) -> dict[str, Any]:
    try:
        payload = response.json()
    except (ValueError, requests.exceptions.JSONDecodeError) as exc:
        preview = (response.text or "<empty response>").strip()[:1500]
        raise AppError(
            "CEEPS returned HTTP 200 but the body is not valid JSON. "
            f"Response: {preview}"
        ) from exc

    if not isinstance(payload, dict):
        raise AppError("CEEPS returned a non-object JSON response.")

    readings = payload.get("meterReadings")
    if not isinstance(readings, list):
        raise AppError("CEEPS response does not contain a valid 'meterReadings' list.")

    return payload


def read_usage_points(uploaded_file) -> list[str]:
    try:
        df = pd.read_excel(
            uploaded_file,
            converters={"Merilna točka": str},
        )
    except Exception as exc:
        raise AppError(f"Could not read the metering-point file: {exc}") from exc

    column = "Merilna točka"
    if column not in df.columns:
        raise AppError(
            f"The metering-point workbook must contain the '{column}' column."
        )

    points = (
        df[column]
        .dropna()
        .astype(str)
        .str.strip()
    )

    points = [
        point for point in points
        if point and point.lower() != "nan"
    ]

    points = list(dict.fromkeys(points))

    if not points:
        raise AppError("No valid metering points were found.")

    return points


def read_reference(uploaded_file) -> dict[str, tuple[int, str, str]]:
    try:
        workbook = pd.ExcelFile(uploaded_file)
    except Exception as exc:
        raise AppError(
            f"Could not open the distribution reference workbook: {exc}"
        ) from exc

    missing_sheets = [
        sheet for sheet in CATEGORIES
        if sheet not in workbook.sheet_names
    ]
    if missing_sheets:
        raise AppError(
            "Reference workbook is missing sheet(s): "
            + ", ".join(missing_sheets)
        )

    lookup: dict[str, tuple[int, str, str]] = {}
    duplicates = set()

    for sheet in CATEGORIES:
        try:
            df = pd.read_excel(workbook, sheet_name=sheet)
        except Exception as exc:
            raise AppError(
                f"Could not read reference sheet '{sheet}': {exc}"
            ) from exc

        missing_columns = REFERENCE_COLUMNS.difference(df.columns)
        if missing_columns:
            raise AppError(
                f"Reference sheet '{sheet}' is missing column(s): "
                + ", ".join(sorted(missing_columns))
            )

        for row in df[
            ["merilna_tocka", "distribucija", "naziv_placnika"]
        ].itertuples(index=False):
            point = str(row.merilna_tocka).strip()

            if not point or point.lower() == "nan":
                continue

            try:
                distribution = int(row.distribucija)
            except (TypeError, ValueError):
                continue

            if distribution not in DISTRIBUTIONS:
                continue

            payer = (
                ""
                if pd.isna(row.naziv_placnika)
                else str(row.naziv_placnika).strip()
            )

            if point in lookup:
                duplicates.add(point)
                continue

            lookup[point] = (
                distribution,
                sheet,
                payer,
            )

    if not lookup:
        raise AppError("No usable mappings were found in the reference workbook.")

    if duplicates:
        st.warning(
            f"{len(duplicates)} duplicate reference mapping(s) were found. "
            "The first occurrence is used."
        )

    return lookup


def extract_series(
    meter_reading: dict[str, Any],
) -> tuple[str, list[pd.Series], bool]:
    point = str(meter_reading.get("usagePoint", "")).strip()
    if not point:
        raise AppError("A meter reading is missing 'usagePoint'.")

    blocks = meter_reading.get("intervalBlocks")
    if not isinstance(blocks, list) or not blocks:
        return point, [], False

    output = []
    quality_flag = False

    for block in blocks:
        if not isinstance(block, dict):
            continue

        readings = block.get("intervalReadings")
        if not isinstance(readings, list) or not readings:
            continue

        timestamps = []
        values = []

        for reading in readings:
            if not isinstance(reading, dict):
                continue

            timestamp = reading.get("timestamp")
            value = reading.get("value")

            if timestamp is None or value is None:
                continue

            qualities = reading.get("readingQualities")
            if isinstance(qualities, list) and qualities:
                quality_flag = True

            timestamps.append(timestamp)
            values.append(value)

        if not timestamps:
            continue

        parsed_ts = pd.to_datetime(
            timestamps,
            errors="coerce",
            utc=True,
        )

        numeric_values = pd.to_numeric(
            pd.Series(values, dtype="object"),
            errors="coerce",
        )

        valid = (~parsed_ts.isna()) & (~numeric_values.isna())

        if not valid.any():
            continue

        index = pd.DatetimeIndex(parsed_ts[valid]).tz_convert(None)
        series = pd.Series(
            numeric_values[valid].to_numpy(),
            index=index,
            dtype="float64",
        )

        if series.index.has_duplicates:
            series = series.groupby(level=0).sum()

        output.append(series)

    return point, output, quality_flag


def add_series(
    storage,
    category: str,
    distribution: int,
    payer: str,
    point: str,
    series_list: list[pd.Series],
) -> None:
    storage[(category, distribution)][(payer, point)].extend(series_list)


def build_dataframe(columns) -> pd.DataFrame:
    series_list = []

    for column_key, parts in columns.items():
        if not parts:
            continue

        series = parts[0] if len(parts) == 1 else pd.concat(parts)

        if series.index.has_duplicates:
            series = series.groupby(level=0).sum()

        series = series.sort_index()
        series.name = column_key
        series_list.append(series)

    if not series_list:
        return pd.DataFrame()

    df = pd.concat(series_list, axis=1)
    df.columns = pd.MultiIndex.from_tuples(
        df.columns,
        names=["Payer", "Metering point"],
    )
    df = df.sort_index()
    df = df.reindex(sorted(df.columns), axis=1)
    df.index.name = "timestamp"
    return df


def sheet_name(category: str, distribution: int) -> str:
    return f"{CATEGORIES[category]}_{DISTRIBUTIONS[distribution]}"[:31]


def build_excel(storage) -> bytes:
    output = BytesIO()
    created_sheets = 0

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        timestamp_format = workbook.add_format(
            {"num_format": "yyyy-mm-dd hh:mm:ss"}
        )

        for category in CATEGORIES:
            for distribution in DISTRIBUTIONS:
                columns = storage.get((category, distribution))
                if not columns:
                    continue

                df = build_dataframe(columns)
                if df.empty:
                    continue

                name = sheet_name(category, distribution)

                df.to_excel(
                    writer,
                    sheet_name=name,
                    merge_cells=True,
                )

                worksheet = writer.sheets[name]
                worksheet.freeze_panes(3, 1)
                worksheet.set_column(0, 0, 20, timestamp_format)
                worksheet.set_column(1, len(df.columns), 16)
                worksheet.set_selection(3, 1, 3, 1)

                created_sheets += 1

        if created_sheets == 0:
            raise AppError(
                "No matched readings were available for the Excel workbook."
            )

    output.seek(0)
    return output.getvalue()


def json_filename(point: str, existing: set[str]) -> str:
    base = "".join(
        char if char.isalnum() or char in ("-", "_", ".") else "_"
        for char in point
    ).strip("._") or "meter_reading"

    filename = f"{base}.json"
    counter = 2

    while filename in existing:
        filename = f"{base}_{counter}.json"
        counter += 1

    existing.add(filename)
    return filename


def build_zip(
    json_files: list[tuple[str, bytes]],
    excel_bytes: bytes,
) -> tuple[bytes, str]:
    date_text = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d")
    excel_name = f"Metered_data_{date_text}.xlsx"
    zip_name = f"Metered_data_{date_text}.zip"

    output = BytesIO()

    with zipfile.ZipFile(
        output,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as archive:
        archive.writestr(excel_name, excel_bytes)

        for filename, content in json_files:
            archive.writestr(f"JSON/{filename}", content)

    output.seek(0)
    return output.getvalue(), zip_name


def run_pipeline(
    usage_points: list[str],
    reference: dict[str, tuple[int, str, str]],
    message_type: str,
    ceeps_id: str,
    start_date,
    end_date,
):
    storage = defaultdict(lambda: defaultdict(list))
    json_files = []
    json_names = set()

    returned = set()
    missing_reference = set()
    empty_readings = set()
    quality_flags = set()
    errors = []

    chunks = math.ceil(len(usage_points) / CHUNK_SIZE)
    progress = st.progress(0)
    status = st.empty()

    for index in range(chunks):
        chunk_number = index + 1
        chunk = usage_points[
            index * CHUNK_SIZE:(index + 1) * CHUNK_SIZE
        ]

        status.info(
            f"Retrieving chunk {chunk_number} of {chunks} "
            f"({len(chunk)} metering points)..."
        )

        try:
            response = retrieve_chunk(
                message_type,
                ceeps_id,
                chunk,
                start_date,
                end_date,
            )

            if response.status_code != 200:
                raise AppError(http_error(response))

            payload = parse_response(response)

        except AppError as exc:
            errors.append(
                {
                    "Chunk": chunk_number,
                    "Error": str(exc),
                }
            )
            progress.progress(chunk_number / chunks)
            continue

        for reading in payload["meterReadings"]:
            if not isinstance(reading, dict):
                continue

            point = str(reading.get("usagePoint", "")).strip()
            if not point:
                continue

            returned.add(point)

            name = json_filename(point, json_names)
            json_files.append(
                (
                    name,
                    json.dumps(
                        reading,
                        ensure_ascii=False,
                        indent=2,
                    ).encode("utf-8"),
                )
            )

            try:
                point, series_list, has_quality = extract_series(reading)
            except AppError as exc:
                errors.append(
                    {
                        "Chunk": chunk_number,
                        "Error": f"{point}: {exc}",
                    }
                )
                continue

            if has_quality:
                quality_flags.add(point)

            if not series_list:
                empty_readings.add(point)
                continue

            mapping = reference.get(point)
            if mapping is None:
                missing_reference.add(point)
                continue

            distribution, category, payer = mapping

            add_series(
                storage,
                category,
                distribution,
                payer,
                point,
                series_list,
            )

        progress.progress(chunk_number / chunks)

    if not json_files:
        raise AppError(
            "CEEPS did not return any meter-reading JSON files."
        )

    status.info("Building the consolidated Excel workbook and final ZIP package...")

    excel_bytes = build_excel(storage)
    zip_bytes, zip_name = build_zip(
        json_files,
        excel_bytes,
    )

    not_returned = sorted(set(usage_points) - returned)

    progress.progress(1.0)
    status.empty()

    return zip_bytes, zip_name, {
        "requested": len(usage_points),
        "returned": len(returned),
        "json_files": len(json_files),
        "not_returned": not_returned,
        "missing_reference": sorted(missing_reference),
        "empty_readings": sorted(empty_readings),
        "quality_flags": sorted(quality_flags),
        "errors": errors,
    }


def render_summary(summary: dict[str, Any]) -> None:
    st.subheader("Processing summary")

    c1, c2, c3 = st.columns(3)
    c1.metric("Requested points", summary["requested"])
    c2.metric("Returned points", summary["returned"])
    c3.metric("JSON files", summary["json_files"])

    sections = [
        (
            "Requested points not returned by CEEPS",
            summary["not_returned"],
        ),
        (
            "Returned points missing from the distribution reference",
            summary["missing_reference"],
        ),
        (
            "Metering points with reading-quality flags",
            summary["quality_flags"],
        ),
        (
            "Metering points with no usable interval readings",
            summary["empty_readings"],
        ),
    ]

    for title, values in sections:
        if values:
            with st.expander(f"{title} ({len(values)})"):
                st.dataframe(
                    pd.DataFrame({"Metering point": values}),
                    use_container_width=True,
                    hide_index=True,
                )

    if summary["errors"]:
        with st.expander(
            f"Retrieval / processing errors ({len(summary['errors'])})",
            expanded=True,
        ):
            st.dataframe(
                pd.DataFrame(summary["errors"]),
                use_container_width=True,
                hide_index=True,
            )


def main():
    st.set_page_config(
        page_title="CEEPS Metered Data Processor",
        layout="wide",
    )

    st.title("CEEPS Metered Data Processor")
    st.caption(
        "Retrieve CEEPS meter readings, match them to distribution metadata, "
        "build one consolidated Excel workbook, and package the raw JSON files "
        "and Excel output into a single ZIP archive."
    )

    with st.sidebar:
        st.subheader("Retrieval settings")

        ceeps_id = st.selectbox(
            "CEEPS identity",
            ("NME", "SFA"),
        )

        message_type = st.selectbox(
            "Meter-reading type",
            (
                "Daily 15 minute",
                "Monthly 15 minute",
                "Specify date",
            ),
        )

        start_date = ""
        end_date = ""

        if message_type == "Specify date":
            start_date = st.date_input("Start date")
            end_date = st.date_input("End date")

            if start_date > end_date:
                st.error("Start date cannot be later than end date.")

        st.caption(
            "CEEPS requests are sent in batches of 50 metering points "
            "with timeout, retry, HTTP, and JSON error handling."
        )

    left, right = st.columns(2)

    with left:
        meter_file = st.file_uploader(
            "1. Upload metering-point list",
            type=["xlsx"],
            help="Required column: 'Merilna točka'.",
        )

    with right:
        reference_file = st.file_uploader(
            "2. Upload distribution reference workbook",
            type=["xlsx"],
            help=(
                "Required sheets: dobava, odkup, obratovalna_podpora. "
                "Required columns: merilna_tocka, distribucija, naziv_placnika."
            ),
        )

    if meter_file is None or reference_file is None:
        st.info("Upload both Excel files to continue.")
        return

    if message_type == "Specify date" and start_date > end_date:
        return

    try:
        usage_points = read_usage_points(meter_file)
        reference = read_reference(reference_file)
    except AppError as exc:
        st.error(str(exc))
        return

    found = sum(point in reference for point in usage_points)

    c1, c2, c3 = st.columns(3)
    c1.metric("Unique metering points", len(usage_points))
    c2.metric("Found in reference", found)
    c3.metric("Missing from reference", len(usage_points) - found)

    st.info(
        "The final ZIP contains a JSON folder with the retrieved meter-reading "
        "files and one Metered_data_YYYY-MM-DD.xlsx workbook in the ZIP root."
    )

    if st.button(
        "Retrieve, match and build ZIP",
        type="primary",
        use_container_width=True,
    ):
        st.session_state.pop("metered_zip", None)
        st.session_state.pop("metered_zip_name", None)
        st.session_state.pop("metered_summary", None)

        try:
            zip_bytes, zip_name, summary = run_pipeline(
                usage_points,
                reference,
                message_type,
                ceeps_id,
                start_date,
                end_date,
            )

            st.session_state["metered_zip"] = zip_bytes
            st.session_state["metered_zip_name"] = zip_name
            st.session_state["metered_summary"] = summary

            st.success(
                "Retrieval and processing completed successfully."
            )

        except MemoryError:
            st.error(
                "The server ran out of memory while building the output. "
                "Try a smaller batch or increase the Streamlit instance memory."
            )
        except AppError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"Unexpected error: {exc}")

    summary = st.session_state.get("metered_summary")
    if summary:
        render_summary(summary)

    zip_bytes = st.session_state.get("metered_zip")
    zip_name = st.session_state.get("metered_zip_name")

    if zip_bytes and zip_name:
        st.download_button(
            "Download complete ZIP package",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )


main()
