import tkinter as tk
from tkinter import messagebox
from workweek import CalendarBuilder


class ShiftConfigurationWizard:
    def __init__(self, master):
        self.employee_text = None
        self.master = master
        master.title("Shift Allocation Wizard")

        # Data storage
        self.employee_list = []
        self.shift_config = {}  # Will store {'Morning': 2, 'Swing': 2, 'Night': 1, ...}

        # Shift variables for Step 2
        self.shifts_available = {
            'Morning': tk.IntVar(value=1),
            'Swing': tk.IntVar(value=1),
            'Night': tk.IntVar(value=0)
        }
        self.team_size_entries = {}

        self.current_frame = None
        self.show_employee_input()

    def clear_frame(self):
        """Destroys the current frame to prepare for the next step."""
        if self.current_frame:
            self.current_frame.destroy()

    # --- Step 1: Employee Input (Same as before) ---
    def show_employee_input(self):
        self.clear_frame()
        self.current_frame = tk.Frame(self.master, padx=10, pady=10)
        self.current_frame.pack()

        tk.Label(self.current_frame, text="Step 1: Enter Employee Names (One per line)",
                 font=('Arial', 12, 'bold')).pack(pady=5)

        self.employee_text = tk.Text(self.current_frame, height=10, width=40)
        self.employee_text.pack(pady=5)
        # Placeholder name
        print("Generating placeholder names to be fulfilled by the user")
        self.employee_text.insert("1.0",
                                  "Ahmad Lutfi bin Razak\n"
                                  "Tan Su Ling\n"
                                  "Siva Kumar a/l Ganesan\n"
                                  "Gopal Singh\n"
                                  "Nurul Izzah binti Zulkifli\n"
                                  "Wong Chee Kin\n"
                                  "Priya Devi a/p Murugan\n"
                                  "Harjeet Kaur\n"
                                  "Zul Arifin\n", )  # Prefill for testing

        print(self.employee_text.get('1.0', tk.END))

        tk.Button(self.current_frame, text="Next: Configure Shifts", command=self.process_employee_input).pack(pady=10)

    def process_employee_input(self):
        employee_input = self.employee_text.get("1.0", tk.END).strip()
        self.employee_list = [name.strip() for name in employee_input.split('\n') if name.strip()]

        if not self.employee_list:
            messagebox.showerror("Input Error", "Please enter at least one employee name.")
            return

        self.show_shift_configuration()

    # --- Step 2: Shift Configuration and Allocation ---
    def show_shift_configuration(self):
        self.clear_frame()
        self.current_frame = tk.Frame(self.master, padx=10, pady=10)
        self.current_frame.pack()

        tk.Label(self.current_frame, text="Step 2: Define Shift Configuration",
                 font=('Arial', 12, 'bold')).grid(row=0,
                                                  column=0,
                                                  columnspan=3,
                                                  pady=10)

        # Shift Labels and Input Entries
        shifts = ['Morning', 'Swing', 'Night']
        for i, shift_name in enumerate(shifts):
            row_num = i + 1

            # Checkbox to select the shift
            cb = tk.Checkbutton(self.current_frame, text=f"Enable {shift_name}",
                                variable=self.shifts_available[shift_name], state=tk.DISABLED)
            cb.grid(row=row_num, column=0, sticky="w", padx=5, pady=5)

            # Label for allocation input
            tk.Label(self.current_frame, text="Employees Req:").grid(row=row_num, column=1, sticky="e", padx=5)

            # Allocation Entry Box
            # Replaced Entry to Label
            entry = tk.Label(self.current_frame, width=5)
            entry.grid(row=row_num, column=2, padx=5, pady=5)
            self.team_size_entries[shift_name] = entry

            # Set initial values based on common scenario
            if shift_name in ['Morning', 'Swing']:
                entry['text'] = str(int(len(self.employee_list) / len(shifts)))
                tk.Label(self.current_frame, text="(Default: 2)").grid(row=row_num, column=3, sticky="w")
            elif shift_name == 'Night':
                # TODO additional Night shift
                entry['text'] = "1"
                tk.Label(self.current_frame, text="(Min: 1, if enabled)").grid(row=row_num, column=3, sticky="w")

        # Navigation buttons
        row_num = len(shifts) + 1
        tk.Button(self.current_frame, text="Finish: Generate Schedule",
                  command=self.finish_wizard).grid(row=row_num, column=0, columnspan=3, pady=15)
        tk.Button(self.current_frame, text="Back",
                  command=self.show_employee_input).grid(row=row_num + 1, column=0,
                                                         columnspan=3, pady=5)

    def finish_wizard(self):
        self.shift_config = {}
        total_required_employees = 0

        # Validate and collect final data
        try:
            for shift_name, var in self.shifts_available.items():
                if var.get() == 1:  # Check if the shift is enabled
                    team_size = int(self.team_size_entries[shift_name]['text'])

                    # Enforce Night Shift Rule
                    if shift_name == 'Night' and team_size < 1:
                        messagebox.showerror("Validation Error",
                                             "Night shift must have a minimum of 1 employee allocated (Employee E).")
                        return

                    if team_size <= 0:
                        raise ValueError(f"Team size for {shift_name} must be greater than zero.")

                    self.shift_config[shift_name] = team_size
                    total_required_employees += team_size

        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid input for team size: {e}")
            return

        if not self.shift_config:
            messagebox.showerror("Input Error", "You must enable and configure at least one shift.")
            return

        # --- FINAL ACTION ---
        # Data available for your Excel script:
        # self.employee_list: ['A', 'B', 'C', 'D', 'E', 'F']
        # self.shift_config: {'Morning': 2, 'Swing': 2, 'Night': 1} (or similar)

        show_timed_message(root,"Success",
                            f"Schedule data collected!\n\n"
                            f"Total Employees: {len(self.employee_list)}\n"
                            f"Active Shifts: {', '.join(self.shift_config.keys())}\n"
                            f"Total Team Slots Required (for one day): {total_required_employees}\n"
                            f"Now call your Excel generation script here.")

def show_timed_message(parent, title, message, timeout_ms=3000):
    """
    Displays a custom message box that automatically closes after a specified timeout.

    Args:
        parent: The parent Tkinter window (e.g., root).
        title: The title of the message box.
        message: The message to display.
        timeout_ms: The time in milliseconds after which the message box will close.
    """
    popup = tk.Toplevel(parent)
    popup.title(title)
    popup.transient(parent)  # Makes the popup appear on top of the parent
    popup.grab_set()       # Prevents interaction with other windows until popup closes

    # Center the popup relative to the parent window (optional)
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()

    popup_width = 300  # Approximate width for the popup
    popup_height = 250 # Approximate height for the popup

    x = parent_x + (parent_width // 2) - (popup_width // 2)
    y = parent_y + (parent_height // 2) - (popup_height // 2)
    popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

    label = tk.Label(popup, text=message, justify="left", anchor="nw", wraplength=popup_width - 20, bg="white")
    label.pack(expand=True, fill="both", padx=10, pady=10)

    # Schedule the popup to terminate itself after the timeout
    # popup.after(timeout_ms, popup.destroy)
    popup.after(timeout_ms, parent.destroy)

    # Make sure the popup is closed properly when the user closes the parent window
    parent.protocol("WM_DELETE_WINDOW", lambda: on_closing(parent, popup))

    # Wait for the popup to close before returning control to the main window
    parent.wait_window(popup)

def on_closing(parent, popup):
    """Handles proper closing of both parent and popup."""
    popup.destroy()
    parent.destroy()

# --- Main Application Loop ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ShiftConfigurationWizard(root)
    root.mainloop()

    builder = CalendarBuilder(app.shift_config, app.employee_list)
    # TODO hardcoded year
    builder.build_iso_calendar(2025)

# TODO hardcoded year
# TODO additional Night shift
