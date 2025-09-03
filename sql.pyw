import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import os

class ExcelSQLConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel => SQL Dönüştürücü - 2025 Kemal Lokman")
        self.geometry("1000x650")
        self.configure(bg="navy")

        tk.Label(self, text="Excel => SQL Dönüştürücü", font=("Arial", 22, "bold"), fg="white", bg="navy").pack(pady=10)
        tk.Label(self, text="2025 Kemal Lokman", font=("Arial", 14), fg="white", bg="navy").pack(pady=5)

        # Butonlar frame
        button_frame = tk.Frame(self, bg="navy")
        button_frame.pack(pady=10)
        tk.Button(button_frame, text="Tek Dosya Seç", font=("Arial", 14, "bold"),
                  bg="white", fg="black", padx=15, pady=8, command=self.add_single_excel).pack(side="left", padx=5)
        tk.Button(button_frame, text="Toplu Dosya Seç", font=("Arial", 14, "bold"),
                  bg="white", fg="black", padx=15, pady=8, command=self.add_multiple_excels).pack(side="left", padx=5)
        tk.Button(button_frame, text="Hedef Klasör Seç", font=("Arial", 14, "bold"),
                  bg="white", fg="black", padx=15, pady=8, command=self.select_target_folder).pack(side="left", padx=5)
        tk.Button(button_frame, text="SQL Dosyalarını Oluştur", font=("Arial", 14, "bold"),
                  bg="white", fg="black", padx=15, pady=8, command=self.convert).pack(side="left", padx=5)

        tk.Label(self, text="Veritabanı Türü:", fg="white", bg="navy", font=("Arial", 12)).pack(pady=5)
        self.db_combo = ttk.Combobox(self, values=["MySQL", "MSSQL"], state="readonly", font=("Arial", 12))
        self.db_combo.current(0)
        self.db_combo.pack(pady=5)

        # Listboxlar
        frame = tk.Frame(self, bg="navy")
        frame.pack(pady=10, padx=10, fill="both", expand=True)
        left_frame = tk.Frame(frame, bg="navy")
        left_frame.pack(side="left", fill="both", expand=True, padx=5)
        tk.Label(left_frame, text="Seçilen Excel Dosyaları:", fg="white", bg="navy", font=("Arial", 12)).pack()
        self.listbox_excel = tk.Listbox(left_frame, width=40, height=20, font=("Arial", 11))
        self.listbox_excel.pack(fill="both", expand=True, pady=5)

        right_frame = tk.Frame(frame, bg="navy")
        right_frame.pack(side="right", fill="both", expand=True, padx=5)
        tk.Label(right_frame, text="Oluşturulan SQL Dosyaları:", fg="white", bg="navy", font=("Arial", 12)).pack()
        self.listbox_sql = tk.Listbox(right_frame, width=40, height=20, font=("Arial", 11))
        self.listbox_sql.pack(fill="both", expand=True, pady=5)

        # Son mesaj etiketi
        self.status_label = tk.Label(self, text="", fg="white", bg="navy", font=("Arial", 12))
        self.status_label.pack(pady=5)

        self.excel_files = []
        self.target_folder = None

    def add_single_excel(self):
        path = filedialog.askopenfilename(title="Excel dosyası seç", filetypes=[("Excel dosyaları", "*.xlsx *.xls")])
        if path and path not in self.excel_files:
            self.excel_files.append(path)
            self.listbox_excel.insert(tk.END, os.path.basename(path))

    def add_multiple_excels(self):
        paths = filedialog.askopenfilenames(title="Birden fazla Excel dosyası seç", filetypes=[("Excel dosyaları", "*.xlsx *.xls")])
        for path in paths:
            if path not in self.excel_files:
                self.excel_files.append(path)
                self.listbox_excel.insert(tk.END, os.path.basename(path))

    def select_target_folder(self):
        folder = filedialog.askdirectory(title="SQL dosyalarının kaydedileceği klasörü seçin")
        if folder:
            self.target_folder = folder
            self.status_label.config(text=f"Hedef klasör seçildi: {self.target_folder}")

    def convert(self):
        if not self.excel_files:
            self.status_label.config(text="Lütfen en az bir Excel dosyası seçin.")
            return
        if not self.target_folder:
            self.status_label.config(text="Lütfen hedef klasörü seçin.")
            return

        db_type = self.db_combo.get()
        self.listbox_sql.delete(0, tk.END)
        self.status_label.config(text="İşlem başladı...")
        self.update_idletasks()

        for file_path in self.excel_files:
            try:
                df = pd.read_excel(file_path)
                table_name = os.path.splitext(os.path.basename(file_path))[0]
                sql_file = os.path.join(self.target_folder, f"{table_name}.sql")

                with open(sql_file, "w", encoding="utf-8") as f:
                    if db_type == "MySQL":
                        f.write("USE sakila;\n\n")
                        id_col = "ID INT AUTO_INCREMENT PRIMARY KEY"
                        text_type = "VARCHAR(50)"
                        use_backtick = True
                    else:
                        id_col = "ID INT IDENTITY(1,1) PRIMARY KEY"
                        text_type = "NVARCHAR(50)"
                        use_backtick = False

                    f.write(f"CREATE TABLE {table_name} (\n")
                    f.write(f"    {id_col},\n")
                    for idx, col in enumerate(df.columns):
                        if pd.api.types.is_datetime64_any_dtype(df[col]):
                            col_type = "DATE"
                        elif pd.api.types.is_numeric_dtype(df[col]):
                            col_type = "FLOAT"
                        else:
                            col_type = text_type
                        col_name = f"`{col}`" if use_backtick else col
                        end_char = "," if idx < len(df.columns)-1 else ""
                        f.write(f"    {col_name} {col_type}{end_char}\n")
                    f.write(");\n\n")

                    for _, row in df.iterrows():
                        values = []
                        for val in row:
                            if pd.isna(val):
                                values.append("NULL")
                            elif isinstance(val, (int, float)):
                                values.append(str(val))
                            elif isinstance(val, pd.Timestamp):
                                values.append(f"'{val.date()}'")
                            else:
                                val = str(val).replace("'", "''")
                                values.append(f"'{val}'")
                        cols_str = ", ".join([f"`{c}`" if use_backtick else c for c in df.columns])
                        f.write(f"INSERT INTO {table_name} ({cols_str}) VALUES ({', '.join(values)});\n")

                self.listbox_sql.insert(tk.END, os.path.basename(sql_file))

            except Exception as e:
                self.status_label.config(text=f"Hata: {file_path} işlenemedi! {str(e)}")

        self.status_label.config(text="Tüm SQL dosyaları başarıyla oluşturuldu!")

if __name__ == "__main__":
    app = ExcelSQLConverter()
    app.mainloop()
