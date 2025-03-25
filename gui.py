import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import tkinter as tk
from tkinter import filedialog, messagebox

# --- 1. 图形界面设置 ---
def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)


def generate_chart():
    try:
        file_path = entry_file_path.get()
        berth_length = int(entry_berth_length.get())

        # --- 2. 加载数据 ---
        data = pd.read_excel(file_path)

        # --- 3. 转换时间列并计算时间偏移 ---
        time_columns = ['ArrivalTime', 'SailStartTime', 'DepartureTime']
        for col in time_columns:
            data[col] = pd.to_datetime(data[col], unit='s')

        # 以最早的 ArrivalTime 作为时间起点
        simulation_start = data['ArrivalTime'].min()
        data['ArrivalOffset'] = (data['ArrivalTime'] - simulation_start).dt.total_seconds() / 3600
        data['SailStartOffset'] = (data['SailStartTime'] - simulation_start).dt.total_seconds() / 3600
        data['DepartureOffset'] = (data['DepartureTime'] - simulation_start).dt.total_seconds() / 3600

        # --- 4. 创建颜色映射 ---
        colors = cm.get_cmap('tab20', len(data))

        # --- 5. 创建图表 ---
        fig, ax = plt.subplots(figsize=(14, 8))

        # 绘制每条船的泊位占用时间
        for idx, row in data.iterrows():
            berth_position = row['BerthPosition']
            vessel_length = row['VesselLength']
            arrival_offset = row['ArrivalOffset']
            sail_start_offset = row['SailStartOffset']
            departure_offset = row['DepartureOffset']

            # 动态计算条形的左侧起始位置
            bar_left = berth_position - vessel_length / 2

            # 绘制等待时间虚线
            if arrival_offset < sail_start_offset:
                ax.plot(
                    [bar_left + vessel_length / 2, bar_left + vessel_length / 2],
                    [arrival_offset, sail_start_offset],
                    color='gray',
                    linestyle='--',
                    linewidth=1.5,
                    label="Waiting Time" if idx == 0 else ""
                )

            # 绘制船舶的占用条形
            ax.barh(
                y=(sail_start_offset + departure_offset) / 2,
                width=vessel_length,
                left=bar_left,
                height=departure_offset - sail_start_offset,
                color=colors(idx),
                edgecolor='black'
            )

            # 在条形中心添加船舶名称和类型
            ax.text(
                berth_position,
                (sail_start_offset + departure_offset) / 2,
                f"{row['VesselObjectName']}\n{row['VesselType']}",
                ha='center', va='center', fontsize=8, color='white', weight='bold'
            )

        # --- 6. 设置x轴 & y轴 ---
        ax.set_xlim(0, berth_length)
        ax.set_xticks(range(0, berth_length + 1, 100))
        ax.set_xticklabels([f"{i}m" for i in range(0, berth_length + 1, 100)])
        ax.set_ylim(0, data['DepartureOffset'].max() + 10)
        ax.set_ylabel("Time (hours from Arrival Time)")
        ax.yaxis.set_major_locator(plt.MultipleLocator(8))
        ax.grid(axis="y", which="both", linestyle="--", linewidth=0.5)

        # --- 7. 添加图例、标题、标签 ---
        ax.legend(loc='upper right')
        ax.set_title("Berth Occupancy Gantt Chart with Waiting Time")
        ax.set_xlabel("Berth Position (m)")

        # --- 8. 显示图表 ---
        plt.tight_layout()
        plt.show()

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong: {e}")


# --- 9. 设置主窗口 ---
app = tk.Tk()
app.title("Berth Occupancy Gantt Chart Tool")
app.geometry("500x300")

# --- 10. 创建界面部件 ---
tk.Label(app, text="Select Excel File:").pack(pady=5)
entry_file_path = tk.Entry(app, width=50)
entry_file_path.pack(pady=5)
tk.Button(app, text="Browse", command=load_file).pack(pady=5)

tk.Label(app, text="Enter Berth Length (m):").pack(pady=5)
entry_berth_length = tk.Entry(app, width=20)
entry_berth_length.insert(0, "1200")
entry_berth_length.pack(pady=5)

tk.Button(app, text="Generate Chart", command=generate_chart, bg="lightgreen").pack(pady=20)

# --- 11. 启动主窗口循环 ---
app.mainloop()
