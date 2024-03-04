import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class RenewalRateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Renewal Rate Calculator")
        self.root.geometry("1200x800")
        
        # Layout Configuration
        left_frame = tk.Frame(root)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        right_frame = tk.Frame(root)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.upload_button = tk.Button(left_frame, text="Upload Excel File", command=self.upload_file)
        self.upload_button.pack(pady=20)

        self.apply_button = tk.Button(left_frame, text="Apply Changes", command=self.apply_changes, state=tk.DISABLED)
        self.apply_button.pack(pady=5)
        
        self.export_button = tk.Button(left_frame, text="Export to Excel", command=self.export_to_excel, state=tk.DISABLED)
        self.export_button.pack(pady=5)

        
        self.figure, self.ax1 = plt.subplots(figsize=(6, 6))
        self.canvas = FigureCanvasTkAgg(self.figure, left_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Adjust columns to include "Used" and set their width
        self.tree = ttk.Treeview(right_frame, columns=("Company Name", "Current Spend", "Used", "% Growth", "Growth Amount", "New Total"), show="headings", height=20)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Example of setting column width and alignment
        self.tree.column("Company Name", width=200, anchor='center')
        self.tree.column("Current Spend", width=100, anchor='center')
        self.tree.column("Used", width=100, anchor='center')
        self.tree.column("% Growth", width=80, anchor='center')
        self.tree.column("Growth Amount", width=100, anchor='center')
        self.tree.column("New Total", width=100, anchor='center')

        # Configuring column headings
        self.tree.heading("Company Name", text="Company Name")
        self.tree.heading("Current Spend", text="Current Spend ($)")
        self.tree.heading("Used", text="Used")
        self.tree.heading("% Growth", text="% Growth")
        self.tree.heading("Growth Amount", text="Growth Amount ($)")
        self.tree.heading("New Total", text="New Total ($)")

        # Adding a scrollbar
        scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        style = ttk.Style()
        style.configure("Treeview", font=('Helvetica', 12), rowheight=30)
        style.configure("Treeview.Heading", font=('Helvetica', 14, 'bold'))  # Adjust the font for headings

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        self.data = None
        self.selected_companies = []

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.data = pd.read_excel(file_path)
            # Check for required columns, now including "Used"
            if not {'Company Name', 'Customer Value', 'Used'}.issubset(self.data.columns):
                messagebox.showerror("Error", "Uploaded file must contain 'Company Name', 'Customer Value', and 'Used' columns.")
                return
            self.original_total = self.data['Customer Value'].sum()
            self.data['Cancelled'] = False  # Add a new column to track cancellations

            # Call calculate_renewals here to apply initial growth calculations
            self.calculate_renewals()

            self.update_charts()
            self.apply_button['state'] = tk.NORMAL
            self.export_button['state'] = tk.NORMAL
            self.populate_treeview()

    def update_charts(self):
        if self.data is not None:
            self.ax1.clear()
            company_shares = self.data.groupby("Company Name")["Customer Value"].sum()
            self.ax1.pie(company_shares, labels=company_shares.index, autopct='%1.1f%%', startangle=140)
            self.ax1.set_title('Company Shares')
            self.canvas.draw()

    def populate_treeview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in self.data.iterrows():
            # Check if 'Growth Rate' is a number and format accordingly
            if pd.notnull(row.get('Growth Rate')) and isinstance(row['Growth Rate'], (float, int)):
                growth_rate = f"{row['Growth Rate']:.2%}"
            else:
                growth_rate = "N/A"
            
            # Apply similar logic for 'Growth Amount' and 'New Total'
            growth_amount = f"{row.get('Growth Amount', 'N/A'):.2f}" if pd.notnull(row.get('Growth Amount')) else "N/A"
            new_total = f"{row.get('New Total', 'N/A'):.2f}" if pd.notnull(row.get('New Total')) else "N/A"
            
            if row.get('Cancelled', False):
                self.tree.insert("", tk.END, values=(row["Company Name"], row["Customer Value"], row["Used"], growth_rate, growth_amount, new_total), tags=('cancelled',))
            else:
                self.tree.insert("", tk.END, values=(row["Company Name"], row["Customer Value"], row["Used"], growth_rate, growth_amount, new_total))
        self.tree.tag_configure('cancelled', background='red')


    def apply_changes(self):
        # Mark companies as cancelled or not based on the current selection in the Treeview.
        selected_items = self.tree.selection()
        selected_companies = [self.tree.item(item)['values'][0] for item in selected_items]

        # Update the 'Cancelled' status based on selection.
        self.data['Cancelled'] = self.data['Company Name'].isin(selected_companies)

        # Recalculate growth metrics to reflect the current state, including cancellations.
        self.calculate_renewals()

        # Update the charts and Treeview to reflect the new calculations.
        self.update_charts()
        self.populate_treeview()
        messagebox.showinfo("Info", "Changes Applied Successfully")

    def calculate_renewals(self):
        # Apply initial growth calculations to all companies.
        self.apply_growth()

        # Adjust growth calculations based on the current state, including any cancellations.
        self.adjust_growth_for_cancellations()

        # Update the Treeview with the new calculations.
        self.populate_treeview()

    def apply_growth(self):
        # Calculate the average usage across all companies.
        average_usage = self.data['Used'].mean()

        # Define the min and max growth rates.
        min_growth_rate, max_growth_rate = 0.03, 0.08  # 3% to 8%
        
        # Normalize company usage around the average and scale growth rates accordingly.
        # This example assumes 'Used' values are always positive. Adjust logic as needed for your data.
        self.data['Usage Ratio'] = self.data['Used'] / average_usage
        # Scale the growth rate based on the usage ratio. This is a simplistic scaling approach.
        # You might need a more sophisticated formula depending on your distribution of 'Used' values.
        self.data['Growth Rate'] = self.data['Usage Ratio'].apply(lambda x: min_growth_rate + (x * (max_growth_rate - min_growth_rate)))
        # Ensure growth rate does not exceed specified min/max boundaries.
        self.data['Growth Rate'] = self.data['Growth Rate'].clip(lower=min_growth_rate, upper=max_growth_rate)
        
        # Calculate growth amount and new total based on the proportional growth rate.
        self.data['Growth Amount'] = self.data['Customer Value'] * self.data['Growth Rate']
        self.data['New Total'] = self.data['Customer Value'] + self.data['Growth Amount']

    def adjust_growth_for_cancellations(self):
        # Calculate the total loss due to cancellations.
        cancelled_total = self.data.loc[self.data['Cancelled'], 'Customer Value'].sum()
        growth_total = self.data.loc[~self.data['Cancelled'], 'Growth Amount'].sum()

        # Determine if the growth covers the loss from cancellations by an additional 2%.
        required_growth_to_cover_loss = cancelled_total * 1.02  # 2% more than the loss.
        growth_shortfall = required_growth_to_cover_loss - growth_total

        if growth_shortfall > 0:
            # If growth does not cover the loss, adjust growth proportionally.
            non_cancelled_total = self.data.loc[~self.data['Cancelled'], 'Customer Value'].sum()

            # Calculate additional growth needed for each company to cover the shortfall and distribute it.
            self.data.loc[~self.data['Cancelled'], 'Additional Growth Needed'] = (self.data['Customer Value'] / non_cancelled_total) * growth_shortfall
            self.data.loc[~self.data['Cancelled'], 'Growth Amount'] += self.data.loc[~self.data['Cancelled'], 'Additional Growth Needed']
            self.data.loc[~self.data['Cancelled'], 'New Total'] = self.data['Customer Value'] + self.data['Growth Amount']
            
            # Update growth rate based on new growth amount.
            self.data.loc[~self.data['Cancelled'], 'Growth Rate'] = self.data.loc[~self.data['Cancelled'], 'Growth Amount'] / self.data.loc[~self.data['Cancelled'], 'Customer Value']

    def export_to_excel(self):
        # Define the file name and path to save the Excel file.
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not filepath:
            # User cancelled; don't proceed.
            return

        try:
            # Select columns to export and save to an Excel file.
            columns_to_export = ['Company Name', 'Customer Value', 'Used', 'Growth Rate', 'Growth Amount', 'New Total', 'Cancelled']
            self.data[columns_to_export].to_excel(filepath, index=False)

            messagebox.showinfo("Export Successful", f"Data successfully exported to {filepath}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))


    def calculate_renewals(self):
        self.apply_growth()
        self.adjust_growth_for_cancellations()
        self.populate_treeview()  # Update Treeview with new calculations

root = tk.Tk()
app = RenewalRateApp(root)
root.mainloop()
