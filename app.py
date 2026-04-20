import io
import pandas as pd
from shiny import App, render, ui, reactive
from openpyxl.styles import PatternFill

# Color Constants
PAIR_COLORS = ["#FFF2CC", "#D9EAD3", "#D9EAF7", "#F4CCCC", "#EAD1DC"]
EXCEL_COLORS = ["FFF2CC", "D9EAD3", "D9EAF7", "F4CCCC", "EAD1DC"]

app_ui = ui.page_fluid(
    ui.panel_title("Excel Comparator Lite"),
    ui.layout_sidebar(
        ui.sidebar(
            ui.h5("Source Files"),
            ui.input_file("file1", "Finance / File 1", accept=[".xlsx"], multiple=False),
            ui.output_ui("sheet_sel1"),
            ui.input_numeric("header1", "Header Row Index (File 1)", 0, min=0),
            
            ui.hr(),
            
            ui.input_file("file2", "CPUS / File 2", accept=[".xlsx"], multiple=False),
            ui.output_ui("sheet_sel2"),
            ui.input_numeric("header2", "Header Row Index (File 2)", 0, min=0),
            
            ui.hr(),
            ui.input_radio_buttons("mode", "Filter Mode", {
                "in": "ANY File 2 column is IN File 1",
                "not_in": "ANY File 2 column is NOT IN File 1"
            }),
            ui.input_action_button("compare", "Compare 5 Columns", class_="btn-primary w-100"),
            ui.hr(),
            ui.download_button("download", "Export to Excel", class_="w-100"),
        ),
        ui.navset_card_pill(
            ui.nav_panel("Comparison Results",
                ui.output_ui("report_summary"),
                ui.output_data_frame("result_table")
            ),
            ui.nav_panel("Column Pairing",
                ui.output_ui("pair_selectors")
            )
        )
    )
)

def server(input, output, session):
    
    def read_excel(file_info, sheet, header):
        if file_info is None: return None
        return pd.read_excel(file_info[0]["datapath"], sheet_name=sheet, header=header)

    @reactive.calc
    def df1_meta():
        if input.file1() is None: return None
        xl = pd.ExcelFile(input.file1()[0]["datapath"])
        return xl.sheet_names

    @reactive.calc
    def df2_meta():
        if input.file2() is None: return None
        xl = pd.ExcelFile(input.file2()[0]["datapath"])
        return xl.sheet_names

    @output
    @render.ui
    def sheet_sel1():
        names = df1_meta()
        if not names: return None
        return ui.input_select("sheet1", "Select Sheet (File 1)", choices=names)

    @output
    @render.ui
    def sheet_sel2():
        names = df2_meta()
        if not names: return None
        return ui.input_select("sheet2", "Select Sheet (File 2)", choices=names)

    @reactive.calc
    def get_dfs():
        f1 = read_excel(input.file1(), input.sheet1(), input.header1())
        f2 = read_excel(input.file2(), input.sheet2(), input.header2())
        return f1, f2

    @output
    @render.ui
    def pair_selectors():
        d1, d2 = get_dfs()
        if d1 is None or d2 is None: return ui.p("Upload files to configure pairs.")
        
        cols1 = d1.columns.tolist()
        cols2 = d2.columns.tolist()
        
        selectors = []
        for i in range(5):
            selectors.append(ui.row(
                ui.column(5, ui.input_select(f"p1_{i}", f"File 1 - Pair {i+1}", choices=[""] + cols1)),
                ui.column(5, ui.input_select(f"p2_{i}", f"File 2 - Pair {i+1}", choices=[""] + cols2)),
                ui.column(2, ui.div(f"Color {i+1}", style=f"background-color: {PAIR_COLORS[i]}; padding: 10px; border: 1px solid #ccc; text-align: center; margin-top: 25px;"))
            ))
        return ui.div(*selectors)

    @reactive.calc
    @reactive.event(input.compare)
    def comparison_logic():
        d1, d2 = get_dfs()
        if d1 is None or d2 is None: return None
        
        work_df = d2.copy()
        masks = []
        active_pairs = []

        def normalize(s):
            return s.fillna("").astype(str).str.replace(r"\s+", " ", regex=True).str.strip().str.upper()

        for i in range(5):
            c1 = getattr(input, f"p1_{i}")()
            c2 = getattr(input, f"p2_{i}")()
            if c1 and c2:
                ref = set(normalize(d1[c1]).tolist())
                target = normalize(work_df[c2])
                status_col = f"Pair {i+1} Status"
                work_df[status_col] = target.isin(ref).map({True: "IN", False: "NOT IN"})
                masks.append(work_df[status_col] == "IN" if input.mode() == "in" else work_df[status_col] == "NOT IN")
                active_pairs.append({"c2": c2, "status": status_col, "color": PAIR_COLORS[i], "ex": EXCEL_COLORS[i]})

        if not masks: return None
        
        final_mask = masks[0]
        for m in masks[1:]: final_mask = final_mask | m
        
        res = work_df.loc[final_mask].copy()
        return {"df": res, "active": active_pairs}

    @output
    @render.ui
    def report_summary():
        data = comparison_logic()
        if data is None: return None
        count = len(data["df"])
        return ui.div(
            ui.p(f"Reconciliation complete: {count} records found."),
            ui.div(
                ui.span("PRO TIP: ", style="font-weight: bold; color: #d9534f;"),
                "Need advanced fuzzy matching or automatic mapping? ",
                ui.a("Upgrade to Pro Version", href="#", class_="btn btn-sm btn-outline-danger"),
                style="background: #f8d7da; padding: 10px; border-radius: 5px; margin-bottom: 15px;"
            )
        )

    @output
    @render.data_frame
    def result_table():
        data = comparison_logic()
        if data is None: return None
        df = data["df"]
        
        # Apply styling in browser
        styler = df.style
        for p in data["active"]:
            styler = styler.background_gradient(subset=[p["c2"], p["status"]], cmap="Pastel1", low=0, high=0) # Simplified
            # Shiny's render.data_frame is best with raw df, custom styling requires different handling
        return df

    @render.download(filename="reconciliation_result.xlsx")
    def download():
        data = comparison_logic()
        if data is None: return None
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            data["df"].to_excel(writer, index=False, sheet_name="Comparison")
            workbook = writer.book
            worksheet = writer.sheets["Comparison"]
            
            for p in data["active"]:
                c2_idx = data["df"].columns.get_loc(p["c2"])
                st_idx = data["df"].columns.get_loc(p["status"])
                fmt = workbook.add_format({'bg_color': p['color']})
                
                # Apply to column range (skipping header)
                worksheet.conditional_format(1, c2_idx, len(data["df"]), c2_idx, {'type': 'no_errors', 'format': fmt})
                worksheet.conditional_format(1, st_idx, len(data["df"]), st_idx, {'type': 'no_errors', 'format': fmt})
        
        yield output.getvalue()

app = App(app_ui, server)
