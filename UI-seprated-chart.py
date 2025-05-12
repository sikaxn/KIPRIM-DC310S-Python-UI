import tkinter as tk
from tkinter import ttk
import serial
import serial.tools.list_ports
from collections import deque
import matplotlib
matplotlib.use('TkAgg')  # Ensure matplotlib uses Tkinter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class DC310SGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("KIPRIM DC310S Controller")

        self.serial_conn = None
        self.voltage_history = deque(maxlen=100)
        self.current_history = deque(maxlen=100)

        # Serial Port Selection
        ttk.Label(root, text="Select COM Port:").grid(row=0, column=0, sticky="w")
        self.combobox = ttk.Combobox(root, values=self.list_serial_ports(), state="readonly")
        self.combobox.grid(row=0, column=1, padx=5)
        self.combobox.set(self.get_highest_com_port())
        ttk.Button(root, text="Connect", command=self.connect).grid(row=0, column=2, padx=5)
        ttk.Button(root, text="Disconnect", command=self.disconnect).grid(row=0, column=3, padx=5)
        self.status_indicator = tk.Label(root, text="●", fg="red", font=("Arial", 14))
        self.status_indicator.grid(row=0, column=4, padx=5)

        # Output Controls
        self.on_button = ttk.Button(root, text="Output ON", command=lambda: self.set_output(1), state='disabled')
        self.on_button.grid(row=1, column=0, pady=5)
        self.off_button = ttk.Button(root, text="Output OFF", command=lambda: self.set_output(0), state='disabled')
        self.off_button.grid(row=1, column=1, pady=5)

        # Set Voltage
        ttk.Label(root, text="Set Voltage (V):").grid(row=2, column=0, sticky="w")
        self.voltage_entry = ttk.Entry(root)
        self.voltage_entry.grid(row=2, column=1)
        ttk.Button(root, text="Set", command=self.set_voltage).grid(row=2, column=2)

        # Set Current
        ttk.Label(root, text="Set Current (A):").grid(row=3, column=0, sticky="w")
        self.current_entry = ttk.Entry(root)
        self.current_entry.grid(row=3, column=1)
        ttk.Button(root, text="Set", command=self.set_current).grid(row=3, column=2)

        # Display Labels
        tk.Label(root, text="Voltage (V):").grid(row=4, column=0, sticky="w")
        self.meas_voltage = tk.Label(root, text="---", font=("Arial", 24, "bold"), fg="blue")
        self.meas_voltage.grid(row=4, column=1, sticky="w")

        tk.Label(root, text="Current (A):").grid(row=4, column=2, sticky="w")
        self.meas_current = tk.Label(root, text="---", font=("Arial", 24, "bold"), fg="green")
        self.meas_current.grid(row=4, column=3, sticky="w")

        tk.Label(root, text="Power (W):").grid(row=4, column=4, sticky="w")
        self.meas_power = tk.Label(root, text="---", font=("Arial", 24, "bold"), fg="orange")
        self.meas_power.grid(row=4, column=5, sticky="w")

        # Chart 1: Voltage
        self.fig_voltage, self.ax_voltage = plt.subplots(figsize=(7, 2.5))
        self.canvas_voltage = FigureCanvasTkAgg(self.fig_voltage, master=root)
        self.canvas_voltage.get_tk_widget().grid(row=5, column=0, columnspan=6)
        self.line_voltage, = self.ax_voltage.plot([], [], label="Voltage (V)", color="blue")
        self.ax_voltage.set_ylim(0, 35)
        self.ax_voltage.set_xlim(0, 100)
        self.ax_voltage.set_title("Voltage")
        self.ax_voltage.set_ylabel("V")
        self.ax_voltage.set_xlabel("Samples")
        self.ax_voltage.grid(True)

        # Chart 2: Current
        self.fig_current, self.ax_current = plt.subplots(figsize=(7, 2.5))
        self.canvas_current = FigureCanvasTkAgg(self.fig_current, master=root)
        self.canvas_current.get_tk_widget().grid(row=6, column=0, columnspan=6)
        self.line_current, = self.ax_current.plot([], [], label="Current (A)", color="green")
        self.ax_current.set_ylim(0, 5)
        self.ax_current.set_xlim(0, 100)
        self.ax_current.set_title("Current")
        self.ax_current.set_ylabel("A")
        self.ax_current.set_xlabel("Samples")
        self.ax_current.grid(True)

        # Auto-refresh loop
        self.root.after(1000, self.periodic_refresh)

    def list_serial_ports(self):
        return [port.device for port in serial.tools.list_ports.comports()]

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
        except Exception:
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
        except:
            return None

    def set_output(self, state: int):
        self.send_command(f"output {state}")

    def set_voltage(self):
        value = self.voltage_entry.get()
        try:
            float(value)
            self.send_command(f"voltage {value}")
        except ValueError:
            pass

    def set_current(self):
        value = self.current_entry.get()
        try:
            float(value)
            self.send_command(f"current {value}")
        except ValueError:
            pass

    def load_initial_settings(self):
        v = self.send_command("voltage?")
        c = self.send_command("current?")
        if v:
            self.voltage_entry.delete(0, tk.END)
            self.voltage_entry.insert(0, v)
        if c:
            self.current_entry.delete(0, tk.END)
            self.current_entry.insert(0, c)

    def refresh_measurements(self):
        v_str = self.send_command("measure:voltage?")
        i_str = self.send_command("measure:current?")
        try:
            v = float(v_str)
            self.meas_voltage.config(text=f"{v:.3f}")
            self.voltage_history.append(v)
        except:
            self.meas_voltage.config(text="---")
            self.voltage_history.append(0)

        try:
            i = float(i_str)
            self.meas_current.config(text=f"{i:.3f}")
            self.current_history.append(i)
        except:
            self.meas_current.config(text="---")
            self.current_history.append(0)

        try:
            p = v * i
            self.meas_power.config(text=f"{p:.2f}")
        except:
            self.meas_power.config(text="---")

        self.update_plots()

    def update_plots(self):
        # Voltage chart
        self.line_voltage.set_data(range(len(self.voltage_history)), list(self.voltage_history))
        self.ax_voltage.set_xlim(0, max(100, len(self.voltage_history)))
        self.ax_voltage.relim()
        self.ax_voltage.autoscale_view(True, True, False)
        self.canvas_voltage.draw()

        # Current chart
        self.line_current.set_data(range(len(self.current_history)), list(self.current_history))
        self.ax_current.set_xlim(0, max(100, len(self.current_history)))
        self.ax_current.relim()
        self.ax_current.autoscale_view(True, True, False)
        self.canvas_current.draw()

    def periodic_refresh(self):
        if self.serial_conn and self.serial_conn.is_open:
            self.refresh_measurements()
        self.root.after(1000, self.periodic_refresh)

if __name__ == "__main__":
    root = tk.Tk()
    app = DC310SGUI(root)
    root.mainloop()
