import openpyxl
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES


def process_file(file_path):
    try:
        BOMIdx=3
        QtyIdx=7
        SheetIdx=0
        
        print("Running, going to load workbook...")
        select_button["state"]="disabled"
        #root.update()
        # Load the workbook
        wb = openpyxl.load_workbook(file_path)
        print (wb.sheetnames)
        print("...loaded.")
        #sheet = wb.active
        sheets = wb.sheetnames
        sheet = wb[sheets[SheetIdx]]
        print("...opened sheet ",SheetIdx)
        print("cols: ",sheet.max_column,", rows:",sheet.max_row)
        print("********************")
        for row in sheet.iter_rows(min_col=1, max_col=sheet.max_column, min_row=1, max_row=1):
            for cell in row:
                print(cell.value)
        print("********************")
        print("Now going to process.")
        nestedBeyond2 = True
        runsCounter=0
        while nestedBeyond2:
            nestedFound = 0
            nestedBeyond2 = False
            ancestorCellIndex = 0
            ancestorCellMultiplier = 1
            previousCellMultiplier = 1
            previousBOMlvl = 1
            for row in sheet.iter_rows(min_col=1, max_col=sheet.max_column, min_row=1, max_row=sheet.max_row):
                print(".",end="")
                BOMlvl = row[BOMIdx]
                #print(row[0].value,", ",row[1].value,", ",row[2].value,", ",row[3].value,", ",row[4].value,", ",row[5].value,", ",row[6].value,", ",row[7].value,", ",row[8].value,", ",row[9].value,", ")
                #print("...got a row's bom",BOMlvl.value)
                #print(row[BOMIdx].value," | ",row[QtyIdx].value)
                if BOMlvl.value is not None and isinstance(BOMlvl.value, (int, float)):
                    if BOMlvl.value > previousBOMlvl:  # one below a top
                        ancestorCellIndex = BOMlvl.row - 1
                        ancestorCellMultiplier = previousCellMultiplier
                    elif BOMlvl.value < previousBOMlvl and previousBOMlvl > 2:  # one below a bottom
                        nestedFound += 1
                        nestedBeyond2 = True
                        for subRow in sheet.iter_rows(min_row=ancestorCellIndex + 1, max_row=BOMlvl.row - 1, min_col=1, max_col=sheet.max_column):
                            for subCell in subRow:
                                if subCell.col_idx == (BOMIdx+1):
                                    subCell.value = subCell.value - 1  # BOM LEVEL                        
                                elif subCell.col_idx == (QtyIdx+1):
                                    subCell.value = subCell.value * ancestorCellMultiplier  # QtyPer
                    previousBOMlvl = BOMlvl.value
                    previousCellMultiplier = row[QtyIdx].value
                else: #empty line
                    print("[found empty line]",end="")
                    if previousBOMlvl > 2:  # at a bottom
                        nestedFound += 1
                        nestedBeyond2 = True
                        for subRow in sheet.iter_rows(min_row=ancestorCellIndex + 1, max_row=BOMlvl.row - 1, min_col=1, max_col=sheet.max_column):
                            for subCell in subRow:
                                if subCell.col_idx == (BOMIdx+1):
                                    subCell.value = subCell.value - 1  # BOM LEVEL                        
                                elif subCell.col_idx == (QtyIdx+1):
                                    subCell.value = subCell.value * ancestorCellMultiplier  # QtyPer
                    previousBOMlvl = 1
            print("!")
            #print("IVE FOUND: ", nestedFound)
            modified_file_path = file_path.replace(".xlsx", "_modified"+str(runsCounter)+".xlsx")
            wb.save(modified_file_path)
            runsCounter+=1

        # Save the modified file
        modified_file_path = file_path.replace(".xlsx", "_modified"+".xlsx")
        wb.save(modified_file_path)
        messagebox.showinfo("Success", f"File processed and saved as: {modified_file_path}")
        select_button["state"]="normal"
        print("...done.")
        sys.exit()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def on_drag_and_drop(event):
    # Get the dropped file path
    file_path = event.data.strip()  # Removes extra whitespace or braces
    if file_path.startswith("{") and file_path.endswith("}"):
        file_path = file_path[1:-1]
    process_file(file_path)


def select_file():
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        process_file(file_path)

root = TkinterDnD.Tk()
root.title("Excel Processor with Drag-and-Drop")
root.geometry("400x200")
select_button = tk.Button(root, text="Select or Drop File", command=select_file, padx=20, pady=10)
select_button.pack()

def main():
    # Create a drag-and-drop-enabled Tkinter window
    #root = TkinterDnD.Tk()
    #root.title("Excel Processor with Drag-and-Drop")
    #root.geometry("400x200")

    # Enable drag-and-drop
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drag_and_drop)

    

    root.mainloop()


if __name__ == "__main__":
    main()

'''
#go from top to bottom
    if current cell > previous cell, ancestor = current cell-1; get index
    else if current cell < previous cell and current cell -1 > 2, are at a bottom
        ancestor index + 1-->current cell -1
            val*=ancestor
            BOM--
'''