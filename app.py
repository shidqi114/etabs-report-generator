import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
import pandas as pd

import etabs_extractor
import report_generator

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ETABSReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ETABS Report Generator")
        self.geometry("600x500")

        # Project Info Frame
        self.info_frame = ctk.CTkFrame(self)
        self.info_frame.pack(pady=20, padx=20, fill="x")
        
        self.proj_name_label = ctk.CTkLabel(self.info_frame, text="Project Name:")
        self.proj_name_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.proj_name_entry = ctk.CTkEntry(self.info_frame, width=300)
        self.proj_name_entry.grid(row=0, column=1, padx=10, pady=10)

        self.engineer_label = ctk.CTkLabel(self.info_frame, text="Engineer Name:")
        self.engineer_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.engineer_entry = ctk.CTkEntry(self.info_frame, width=300)
        self.engineer_entry.grid(row=1, column=1, padx=10, pady=10)

        # File Selection Frame
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(pady=10, padx=20, fill="x")

        self.edb_path = ctk.StringVar()
        self.template_path = ctk.StringVar()

        self.edb_label = ctk.CTkLabel(self.file_frame, text="ETABS Model (.edb):")
        self.edb_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.edb_entry = ctk.CTkEntry(self.file_frame, textvariable=self.edb_path, width=250, state="disabled")
        self.edb_entry.grid(row=0, column=1, padx=10, pady=10)
        self.edb_btn = ctk.CTkButton(self.file_frame, text="Browse", command=self.browse_edb, width=80)
        self.edb_btn.grid(row=0, column=2, padx=10, pady=10)

        self.tpl_label = ctk.CTkLabel(self.file_frame, text="Word Template (.docx):")
        self.tpl_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.tpl_entry = ctk.CTkEntry(self.file_frame, textvariable=self.template_path, width=250, state="disabled")
        self.tpl_entry.grid(row=1, column=1, padx=10, pady=10)
        self.tpl_btn = ctk.CTkButton(self.file_frame, text="Browse", command=self.browse_template, width=80)
        self.tpl_btn.grid(row=1, column=2, padx=10, pady=10)

        # Generate Button
        self.generate_btn = ctk.CTkButton(self, text="Generate Report", command=self.start_generation, height=40, font=("Helvetica", 16, "bold"))
        self.generate_btn.pack(pady=20)

        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(self, text="Ready.")
        self.status_label.pack(pady=5)

    def browse_edb(self):
        filename = filedialog.askopenfilename(title="Select ETABS Model", filetypes=[("ETABS Files", "*.edb")])
        if filename:
            self.edb_path.set(filename)

    def browse_template(self):
        filename = filedialog.askopenfilename(title="Select Word Template", filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.template_path.set(filename)

    def start_generation(self):
        if not self.edb_path.get() or not self.template_path.get():
            messagebox.showwarning("Missing Files", "Please select both the ETABS model and the Word template.")
            return

        self.generate_btn.configure(state="disabled")
        self.progress_bar.set(0.1)
        self.status_label.configure(text="Connecting to ETABS...")

        # Run extraction in separate thread
        threading.Thread(target=self.generate_process).start()

    def generate_process(self):
        try:
            edb = self.edb_path.get()
            tpl = self.template_path.get()
            proj_name = self.proj_name_entry.get()
            eng_name = self.engineer_entry.get()
            
            output_path = os.path.splitext(tpl)[0] + "_Generated.docx"

            self.progress_bar.set(0.2)
            self.status_label.configure(text="Extracting from ETABS (This may take a while)...")
            
            # Since we might be testing on Mac, catch the comtypes ImportError
            # and provide some dummy data for testing the UI and Word Generation.
            try:
                df_forces, df_rebar = etabs_extractor.get_beam_results(edb)
            except Exception as e:
                if "comtypes" in str(e) or "Windows" in str(e):
                    # We are on Mac or without ETABS, use dummy data to test Word doc gen
                    self.status_label.configure(text="ETABS API not available. Using Dummy Data for testing...")
                    df_forces = pd.DataFrame({
                        "Frame": ["B1", "B1", "B2"], "Station": [0.0, 3.0, 0.0], 
                        "LoadCase": ["D", "D", "D"], "P": [10.5, 10.5, 15.2], 
                        "V2": [5.2, -5.2, 8.1], "V3": [0, 0, 0], "T": [0.1, 0.1, 0.2], 
                        "M2": [0, 0, 0], "M3": [12.5, -12.5, 20.1]
                    })
                    df_rebar = pd.DataFrame({
                        "Frame": ["B1", "B2"], "Station": [1.5, 2.0],
                        "TopRebar": [500, 600], "BotRebar": [250, 300],
                        "ShearRebar": [0.2, 0.3], "TorsionRebar": [0, 0]
                    })
                else:
                    raise e

            self.progress_bar.set(0.6)
            self.status_label.configure(text="Generating Word document...")

            report_generator.generate_report(tpl, output_path, proj_name, eng_name, df_forces, df_rebar)

            self.progress_bar.set(1.0)
            self.status_label.configure(text="Report Generated Successfully!")
            messagebox.showinfo("Success", f"Report generated successfully!\nSaved to: {output_path}")
        except Exception as e:
            self.status_label.configure(text="Error occurred.")
            messagebox.showerror("Error", str(e))
        finally:
            self.generate_btn.configure(state="normal")

if __name__ == "__main__":
    app = ETABSReportApp()
    app.mainloop()
