import pandas as pd
import json
import sys

def convert_excel_to_json(excel_path, output_path="summary_with_history.json"):
    df = pd.read_excel(excel_path)

    stations = []
    for station_id, group in df.groupby("站別"):
        group_sorted = group.sort_values("日期")
        latest = group_sorted.iloc[-1]

        history = []
        for _, row in group_sorted.iterrows():
            history.append({
                "date": row["日期"].strftime("%Y-%m-%d") if hasattr(row["日期"], "strftime") else str(row["日期"]),
                "accumulated": int(row["上次累積"]),
                "new_issues": int(row["本次新增"]),
                "resolved": int(row["已修正"])
            })

        latest_acc = int(latest["上次累積"])
        latest_new = int(latest["本次新增"])
        latest_resolved = int(latest["已修正"])

        stations.append({
            "id": station_id,
            "name": station_id,
            "last_accumulated": latest_acc,
            "new_issues": latest_new,
            "resolved": latest_resolved,
            "details": f"{station_id} 本次新增 {latest_new} 筆，已修正 {latest_resolved} 筆",
            "history": history
        })

    summary_data = { "stations": stations }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(summary_data, f, ensure_ascii=False, indent=2)

    print(f"✅ 輸出成功：{output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("❗請提供 Excel 檔案路徑，例如：python excel_to_summary_json.py data.xlsx")
    else:
        excel_file = sys.argv[1]
        convert_excel_to_json(excel_file)
