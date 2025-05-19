import json
import os
import tkinter as tk
from tkinter import ttk, simpledialog
import serial
import serial.tools.list_ports
from collections import deque
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

PRESET_FILE = "presets.json"
RESET_SETTINGS_FILE = "reset_settings.json"

class DC310SGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("KIPRIM DC310S Controller")
        self.root.resizable(True, True)

        self.serial_conn = None
        self.voltage_history = deque(maxlen=100)
        self.current_history = deque(maxlen=100)
        self.output_voltage = 0.0
        self.output_power = 0.0
        self.preset_click_stage = 0
        self.elapsed_seconds = 0
        self.energy_ws = 0.0

        self.auto_reset_mode = self.load_reset_settings()
        self.presets = self.load_or_create_presets()

        left_frame = ttk.Frame(root)
        left_frame.grid(row=0, column=0, rowspan=10, padx=5, sticky="ns")
        right_frame = ttk.Frame(root)
        right_frame.grid(row=0, column=1, rowspan=10, padx=5, sticky="ns")

        self.setup_core_controls(left_frame)
        self.setup_presets_panel(right_frame)
        self.root.after(1000, self.periodic_refresh)

    def load_reset_settings(self):
        if not os.path.exists(RESET_SETTINGS_FILE):
            default = {
                "timer": "reset on output off",
                "energy": "reset on output off",
                "all": "no reset"
            }
            with open(RESET_SETTINGS_FILE, "w") as f:
                json.dump(default, f, indent=2)
            return default
        with open(RESET_SETTINGS_FILE, "r") as f:
            return json.load(f)

    def save_reset_settings(self):
        with open(RESET_SETTINGS_FILE, "w") as f:
            json.dump(self.auto_reset_mode, f, indent=2)

    def setup_core_controls(self, frame):
        ttk.Label(frame, text="Select COM Port:").grid(row=0, column=0, sticky="w")
        self.combobox = ttk.Combobox(frame, values=self.list_serial_ports(), state="readonly")
        self.combobox.grid(row=0, column=1)
        self.combobox.set(self.get_highest_com_port())
        ttk.Button(frame, text="Connect", command=self.connect).grid(row=0, column=2)
        ttk.Button(frame, text="Disconnect", command=self.disconnect).grid(row=0, column=3)
        self.status_indicator = tk.Label(frame, text="●", fg="red", font=("Arial", 14))
        self.status_indicator.grid(row=0, column=4)

        self.on_button = tk.Button(frame, text="Output ON", bg="lightcoral", command=lambda: self.set_output(1), state='disabled')
        self.on_button.grid(row=1, column=0, pady=5)
        self.off_button = tk.Button(frame, text="Output OFF", bg="lightblue", command=lambda: self.set_output(0), state='disabled')
        self.off_button.grid(row=1, column=1, pady=5)

        ttk.Label(frame, text="Set Voltage (V):").grid(row=2, column=0, sticky="w")
        self.voltage_entry = ttk.Entry(frame)
        self.voltage_entry.grid(row=2, column=1)
        ttk.Button(frame, text="Set", command=self.set_voltage).grid(row=2, column=2)

        ttk.Label(frame, text="Set Current (A):").grid(row=3, column=0, sticky="w")
        self.current_entry = ttk.Entry(frame)
        self.current_entry.grid(row=3, column=1)
        ttk.Button(frame, text="Set", command=self.set_current).grid(row=3, column=2)

        tk.Label(frame, text="Voltage (V):").grid(row=4, column=0, sticky="w")
        self.meas_voltage = tk.Label(frame, text="---", font=("Arial", 24, "bold"), fg="blue", width=10)
        self.meas_voltage.grid(row=4, column=1)

        tk.Label(frame, text="Current (A):").grid(row=4, column=2, sticky="w")
        self.meas_current = tk.Label(frame, text="---", font=("Arial", 24, "bold"), fg="green", width=10)
        self.meas_current.grid(row=4, column=3)

        tk.Label(frame, text="Power (W):").grid(row=4, column=4, sticky="w")
        self.meas_power = tk.Label(frame, text="---", font=("Arial", 24, "bold"), fg="orange", width=10)
        self.meas_power.grid(row=4, column=5)

        self.fig_voltage, self.ax_voltage = plt.subplots(figsize=(7, 2.5))
        self.canvas_voltage = FigureCanvasTkAgg(self.fig_voltage, master=frame)
        self.canvas_voltage.get_tk_widget().grid(row=5, column=0, columnspan=6)
        self.line_voltage, = self.ax_voltage.plot([], [], color="blue")
        self.ax_voltage.set_ylim(0, 35)
        self.ax_voltage.set_xlim(0, 100)
        self.ax_voltage.set_title("Voltage")
        self.ax_voltage.grid(True)

        self.fig_current, self.ax_current = plt.subplots(figsize=(7, 2.5))
        self.canvas_current = FigureCanvasTkAgg(self.fig_current, master=frame)
        self.canvas_current.get_tk_widget().grid(row=6, column=0, columnspan=6)
        self.line_current, = self.ax_current.plot([], [], color="green")
        self.ax_current.set_ylim(0, 5)
        self.ax_current.set_xlim(0, 100)
        self.ax_current.set_title("Current")
        self.ax_current.grid(True)

    def setup_presets_panel(self, frame):
        ttk.Label(frame, text="Presets").pack()
        self.preset_listbox = tk.Listbox(frame, height=10, width=35)
        self.preset_listbox.pack(pady=5)

        self.preset_display_map = {}
        for name, data in self.presets.items():
            display = f"{name} ({data['voltage']}V, {data['current']}A)"
            self.preset_display_map[display] = name
            self.preset_listbox.insert(tk.END, display)

        self.load_preset_button = tk.Button(frame, text="LOAD", font=("Arial", 14, "bold"),
                                            bg="lightblue", command=self.load_selected_preset)
        self.load_preset_button.pack(pady=10, fill="x")

        self.save_preset_button = tk.Button(frame, text="Save Current Setting",
                                            command=self.save_current_as_preset)
        self.save_preset_button.pack(pady=5, fill="x")

        self.timer_label = tk.Label(frame, text="Time: 00:00:00", font=("Arial", 24, "bold"), fg="purple")
        self.timer_label.pack(pady=5)

        self.energy_label = tk.Label(frame, text="Energy: 0.00 Wh / 0.00 J", font=("Arial", 24, "bold"), fg="brown")
        self.energy_label.pack(pady=5)

        reset_frame = ttk.Frame(frame)
        reset_frame.pack(pady=2, fill="x")

        tk.Button(reset_frame, text="Reset Timer", font=("Arial", 12), command=self.reset_timer).grid(row=0, column=0, sticky="ew")
        tk.Button(reset_frame, text="Reset Energy", font=("Arial", 12), command=self.reset_energy).grid(row=0, column=1, sticky="ew")
        tk.Button(reset_frame, text="Reset All", font=("Arial", 12), command=self.reset_all).grid(row=0, column=2, sticky="ew")

        for idx, key in enumerate(["timer", "energy", "all"]):
            cb = ttk.Combobox(reset_frame, values=[
                "reset on output on", "reset on output off", "no reset"
            ], state="readonly", width=20)
            cb.set(self.auto_reset_mode.get(key, "no reset"))
            cb.grid(row=1, column=idx)
            cb.bind("<<ComboboxSelected>>", lambda e, k=key, cb=cb: self.set_reset_mode(k, cb.get()))

    def set_reset_mode(self, key, mode):
        self.auto_reset_mode[key] = mode
        self.save_reset_settings()

    def reset_timer(self): self.elapsed_seconds = 0
    def reset_energy(self): self.energy_ws = 0.0
    def reset_all(self): self.reset_timer(); self.reset_energy()

    def load_or_create_presets(self):
        if not os.path.exists(PRESET_FILE):
            presets = {
                "USB Power Supply": {"voltage": 5.0, "current": 3.0},
                "Lead Acid Battery": {"voltage": 13.7, "current": 3.0}
            }
            with open(PRESET_FILE, "w") as f:
                json.dump(presets, f, indent=2)
            return presets
        with open(PRESET_FILE, "r") as f:
            return json.load(f)

    def save_presets(self):
        with open(PRESET_FILE, "w") as f:
            json.dump(self.presets, f, indent=2)

    def save_current_as_preset(self):
        name = simpledialog.askstring("Preset Name", "Enter name for this preset:")
        if not name: return
        try:
            voltage = float(self.voltage_entry.get())
            current = float(self.current_entry.get())
            self.presets[name] = {"voltage": voltage, "current": current}
            display = f"{name} ({voltage}V, {current}A)"
            self.preset_display_map[display] = name
            self.preset_listbox.insert(tk.END, display)
            self.save_presets()
        except ValueError: pass

    def list_serial_ports(self): return [port.device for port in serial.tools.list_ports.comports()]
    def get_highest_com_port(self):
        ports = self.list_serial_ports()
        ports_sorted = sorted(ports, key=lambda x: int(x.replace("COM", "")) if x.replace("COM", "").isdigit() else -1)
        return ports_sorted[-1] if ports_sorted else ""

    def connect(self):
        port = self.combobox.get()
        try:
            self.serial_conn = serial.Serial(port, baudrate=115200, timeout=1)
            self.on_button.config(state='normal')
            self.off_button.config(state='normal')
            self.status_indicator.config(fg="green")
            self.load_initial_settings()
        except:
            self.status_indicator.config(fg="red")

    def disconnect(self):
        if self.serial_conn:
            self.serial_conn.close()
            self.serial_conn = None
        self.on_button.config(state='disabled')
        self.off_button.config(state='disabled')
        self.status_indicator.config(fg="red")

    def send_command(self, cmd):
        if not self.serial_conn or not self.serial_conn.is_open:
            return None
        try:
            self.serial_conn.reset_input_buffer()
            self.serial_conn.write((cmd + '\n').encode())
            return self.serial_conn.readline().decode().strip()
        except: return None

    def get_output_state(self): return self.output_voltage > 0.5
    def set_output(self, state: int): self.send_command(f"output {state}")
    def set_voltage(self): self.send_command(f"voltage {self.voltage_entry.get()}")
    def set_current(self): self.send_command(f"current {self.current_entry.get()}")

    def load_initial_settings(self):
        v = self.send_command("voltage?")
        c = self.send_command("current?")
        if v: self.voltage_entry.delete(0, tk.END); self.voltage_entry.insert(0, v)
        if c: self.current_entry.delete(0, tk.END); self.current_entry.insert(0, c)

    def load_selected_preset(self):
        sel = self.preset_listbox.curselection()
        if not sel: return
        display = self.preset_listbox.get(sel[0])
        name = self.preset_display_map.get(display, "")
        preset = self.presets.get(name)
        if not preset: return
        if self.get_output_state() and self.preset_click_stage == 0:
            self.load_preset_button.config(bg="red")
            self.preset_click_stage = 1
            return
        self.voltage_entry.delete(0, tk.END)
        self.voltage_entry.insert(0, str(preset["voltage"]))
        self.current_entry.delete(0, tk.END)
        self.current_entry.insert(0, str(preset["current"]))
        self.set_voltage()
        self.set_current()
        self.load_preset_button.config(bg="lightblue")
        self.preset_click_stage = 0

    def refresh_measurements(self):
        v_str = self.send_command("measure:voltage?")
        i_str = self.send_command("measure:current?")
        try:
            v = float(v_str); self.output_voltage = v
            self.meas_voltage.config(text=f"{v:.3f}")
            self.voltage_history.append(v)
        except:
            self.meas_voltage.config(text="---")
            self.output_voltage = 0
            self.voltage_history.append(0)
        try:
            i = float(i_str)
            self.meas_current.config(text=f"{i:.3f}")
            self.current_history.append(i)
        except:
            self.meas_current.config(text="---")
            self.current_history.append(0)

        power = self.output_voltage * i if v_str and i_str else 0
        self.output_power = power
        self.meas_power.config(text=f"{power:.2f}" if power else "---")

        state = self.get_output_state()
        if state:
            self.elapsed_seconds += 1
            self.energy_ws += self.output_power
            if self.auto_reset_mode["all"] == "reset on output on": self.reset_all()
            if self.auto_reset_mode["timer"] == "reset on output on": self.reset_timer()
            if self.auto_reset_mode["energy"] == "reset on output on": self.reset_energy()
        else:
            if self.auto_reset_mode["all"] == "reset on output off": self.reset_all()
            if self.auto_reset_mode["timer"] == "reset on output off": self.reset_timer()
            if self.auto_reset_mode["energy"] == "reset on output off": self.reset_energy()

        h = self.elapsed_seconds // 3600
        m = (self.elapsed_seconds % 3600) // 60
        s = self.elapsed_seconds % 60
        self.timer_label.config(text=f"Time: {h:02}:{m:02}:{s:02}")
        self.energy_label.config(text=f"Energy: {self.energy_ws / 3600:.2f} Wh / {self.energy_ws:.2f} J")
        self.root.configure(bg="light green" if state else "light gray")
        self.update_plots()

    def update_plots(self):
        self.line_voltage.set_data(range(len(self.voltage_history)), list(self.voltage_history))
        self.ax_voltage.set_xlim(0, max(100, len(self.voltage_history)))
        self.ax_voltage.relim(); self.ax_voltage.autoscale_view(True, True, False); self.canvas_voltage.draw()
        self.line_current.set_data(range(len(self.current_history)), list(self.current_history))
        self.ax_current.set_xlim(0, max(100, len(self.current_history)))
        self.ax_current.relim(); self.ax_current.autoscale_view(True, True, False); self.canvas_current.draw()

    def periodic_refresh(self):
        if self.serial_conn and self.serial_conn.is_open:
            self.refresh_measurements()
        self.root.after(1000, self.periodic_refresh)

if __name__ == "__main__":
    root = tk.Tk()
    app = DC310SGUI(root)
    root.mainloop()
