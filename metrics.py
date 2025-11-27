import statistics
from collections import defaultdict
from datetime import datetime

LABEL_ORDER = [
    "Genera Token",   # Token
    "Purchase",
    "Policy",
    "Generate PDF",
    "Cancel",
]


def to_float(value, default=None):
    try:
        return float(value)
    except Exception:
        return default


def to_int(value, default=0):
    try:
        return int(value)
    except Exception:
        return default


def to_bool_success(value):
    if value is None:
        return False
    v = str(value).strip().lower()
    return v in ("true", "1", "yes", "y")


def percentile(values, p):
    if not values:
        return None
    values = sorted(values)
    k = (len(values) - 1) * (p / 100.0)
    f = int(k)
    c = min(f + 1, len(values) - 1)
    if f == c:
        return values[int(k)]
    d0 = values[f] * (c - k)
    d1 = values[c] * (k - f)
    return d0 + d1


def compute_execution_range_string(rows):
    """
    Calcule la date/heure de début et fin du scénario à partir des timeStamp JMeter.
    Format : 21/11/25 09:42 PM - 21/11/25 10:00 PM
    """
    timestamps = []
    for r in rows:
        ts = r.get("timeStamp")
        if ts is None:
            continue
        try:
            timestamps.append(int(ts))
        except ValueError:
            continue

    if not timestamps:
        return ""

    start_ts = min(timestamps) / 1000.0
    end_ts = max(timestamps) / 1000.0

    start_dt = datetime.fromtimestamp(start_ts)
    end_dt = datetime.fromtimestamp(end_ts)

    fmt = "%d/%m/%y %I:%M %p"
    return f"{start_dt.strftime(fmt)} - {end_dt.strftime(fmt)}"


def compute_recap(rows):
    """
    Retourne une liste de dicts avec :
      Label, Samples, Average (ms), Min (ms), Max (ms), Std Dev (ms),
      Error %, Throughput (/min), Received KB/sec, Sent KB/sec, Avg Bytes
    (90/95/99% calculés mais non stockés)
    """
    labels = {}

    for r in rows:
        label = r.get("label")
        elapsed_raw = r.get("elapsed")
        success_raw = r.get("success")
        ts_raw = r.get("timeStamp")

        if label is None or elapsed_raw is None or ts_raw is None:
            continue

        elapsed = to_float(elapsed_raw)
        if elapsed is None:
            continue

        ts = to_int(ts_raw)
        end_ts = ts + int(elapsed)

        success = to_bool_success(success_raw)
        bytes_val = to_int(r.get("bytes", 0))
        sent_bytes_val = to_int(r.get("sentBytes", 0))

        if label not in labels:
            labels[label] = {
                "times": [],
                "errors": 0,
                "bytes_sum": 0,
                "sent_bytes_sum": 0,
                "first_ts": ts,
                "last_end_ts": end_ts,
            }

        data = labels[label]
        data["times"].append(elapsed)
        if not success:
            data["errors"] += 1
        data["bytes_sum"] += bytes_val
        data["sent_bytes_sum"] += sent_bytes_val
        data["first_ts"] = min(data["first_ts"], ts)
        data["last_end_ts"] = max(data["last_end_ts"], end_ts)

    recap = []

    total_samples_all = 0
    all_times = []
    total_errors_all = 0
    total_bytes_all = 0
    total_sent_bytes_all = 0
    global_first_ts = None
    global_last_end_ts = None

    # ordre des labels pour Word/Excel
    ordered_labels = []
    for lbl in LABEL_ORDER:
        if lbl in labels:
            ordered_labels.append(lbl)
    for lbl in sorted(labels.keys()):
        if lbl not in ordered_labels:
            ordered_labels.append(lbl)

    for label in ordered_labels:
        data = labels[label]
        times = data["times"]
        errors = data["errors"]
        samples = len(times)
        if samples == 0:
            continue

        avg = statistics.mean(times)
        mn = min(times)
        mx = max(times)
        std_dev = statistics.pstdev(times) if samples > 1 else 0.0
        # percentiles calculés mais non stockés
        _ = percentile(times, 90)
        _ = percentile(times, 95)
        _ = percentile(times, 99)
        err_pct = (errors / samples * 100.0) if samples else 0.0

        duration_ms = max(data["last_end_ts"] - data["first_ts"], 1)
        duration_sec = duration_ms / 1000.0
        duration_min = duration_sec / 60.0 if duration_sec > 0 else 0.0

        if duration_min > 0:
            throughput_per_min = samples / duration_min
            recv_kb_per_min = (data["bytes_sum"] / 1024.0) / duration_min
            sent_kb_per_min = (data["sent_bytes_sum"] / 1024.0) / duration_min
        else:
            throughput_per_min = 0.0
            recv_kb_per_min = 0.0
            sent_kb_per_min = 0.0

        avg_bytes = (data["bytes_sum"] / samples) if samples else 0.0

        recap.append({
            "Label": label,
            "Samples": samples,
            "Average (ms)": int(round(avg)),
            "Min (ms)": int(round(mn)),
            "Max (ms)": int(round(mx)),
            "Std Dev (ms)": round(std_dev, 2),
            "Error %": round(err_pct, 2),

            "Throughput (/min)": f"{throughput_per_min:.1f}/min",
            "Received KB/sec": round(recv_kb_per_min, 2),
            "Sent KB/sec": round(sent_kb_per_min, 2),
            "Avg Bytes": round(avg_bytes, 1),
        })

        total_samples_all += samples
        total_errors_all += errors
        all_times.extend(times)
        total_bytes_all += data["bytes_sum"]
        total_sent_bytes_all += data["sent_bytes_sum"]

        if global_first_ts is None or data["first_ts"] < global_first_ts:
            global_first_ts = data["first_ts"]
        if global_last_end_ts is None or data["last_end_ts"] > global_last_end_ts:
            global_last_end_ts = data["last_end_ts"]

    # TOTAL
    if all_times:
        total_avg = statistics.mean(all_times)
        total_min = min(all_times)
        total_max = max(all_times)
        total_std = statistics.pstdev(all_times) if len(all_times) > 1 else 0.0
        _ = percentile(all_times, 90)
        _ = percentile(all_times, 95)
        _ = percentile(all_times, 99)
        total_err_pct = (total_errors_all / total_samples_all * 100.0) if total_samples_all else 0.0

        if global_first_ts is not None and global_last_end_ts is not None:
            duration_ms_all = max(global_last_end_ts - global_first_ts, 1)
            duration_min_all = duration_ms_all / 1000.0 / 60.0
        else:
            duration_min_all = 0.0

        if duration_min_all > 0:
            throughput_all = total_samples_all / duration_min_all
            recv_kb_all = (total_bytes_all / 1024.0) / duration_min_all
            sent_kb_all = (total_sent_bytes_all / 1024.0) / duration_min_all
        else:
            throughput_all = 0.0
            recv_kb_all = 0.0
            sent_kb_all = 0.0

        avg_bytes_all = (total_bytes_all / total_samples_all) if total_samples_all else 0.0

        recap.append({
            "Label": "TOTAL",
            "Samples": total_samples_all,
            "Average (ms)": int(round(total_avg)),
            "Min (ms)": int(round(total_min)),
            "Max (ms)": int(round(total_max)),
            "Std Dev (ms)": round(total_std, 2),
            "Error %": round(total_err_pct, 2),

            "Throughput (/min)": f"{throughput_all:.1f}/min",
            "Received KB/sec": round(recv_kb_all, 2),
            "Sent KB/sec": round(sent_kb_all, 2),
            "Avg Bytes": round(avg_bytes_all, 1),
        })

    return recap
