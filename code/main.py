import os
import time
from datetime import datetime
from io import BytesIO
import input_file
import numpy as np
import worker
import AP
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, \
    QComboBox, QMessageBox, QLineEdit, QDialog
from PySide6.QtGui import QFont
from xhtml2pdf import pisa
from weasyprint import HTML
import copy

# !/usr/bin/env python3


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


# /Users/alexrhe/Documents/AP_AcousticVision_v5_m.xlsx
# /Users/alexrhe/Documents/Worker_AcousticVision.xlsx

# /Users/alexrhe/Documents/Arbeitsplan_HealthStateEngines_V3.xlsx
# /Users/alexrhe/Documents/Worker_HealthStateEngines.xlsx

def get_german_month(english_month):
    months = {
        "January": "Januar",
        "February": "Februar",
        "March": "März",
        "April": "April",
        "May": "Mai",
        "June": "Juni",
        "July": "Juli",
        "August": "August",
        "September": "September",
        "October": "Oktober",
        "November": "November",
        "December": "Dezember"
    }
    return months.get(english_month, "Invalid month")

month_map = {
    'January': 1, 'February': 2, 'March': 3, 'April': 4,
    'May': 5, 'June': 6, 'July': 7, 'August': 8,
    'September': 9, 'October': 10, 'November': 11, 'December': 12
}

# Só para ordenação: converte '3.1' → (3, 1), mas sem modificar o valor original
def ap_id_sort_key(ap_id_str):
    return tuple(map(int, ap_id_str.split('.')))


def calculate_proj_cost(years, hours_year_work_every_one):
    # Add rows for sum worker data
    cost_project = 0
    for i in range(len(years)):
        for j in range(len(worker.list_of_workers)):
            cost_project += round(float(hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]) * \
                            worker.list_of_workers[j].salary,2)
    return cost_project



def show_popup(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setWindowTitle("Information")
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec()


def show_popup_error(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)  # Set icon to Critical for error messages
    msg.setWindowTitle("Error")
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec()


class ExcelReaderApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Files Reader")
        self.setGeometry(100, 100, 800, 300)

        # Create a central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # Labels to display selected file paths
        self.ap_file_label = QLabel("No AP file selected.")
        self.ap_file_label.setObjectName("ap_label")  # Setting object name for styling
        self.ap_file_label.setFont(QFont("Arial", 12))
        layout.addWidget(self.ap_file_label)

        self.worker_file_label = QLabel("No Worker file selected.")
        self.worker_file_label.setObjectName("worker_label")  # Setting object name for styling
        self.worker_file_label.setFont(QFont("Arial", 12))
        layout.addWidget(self.worker_file_label)

        # Buttons to select Excel files
        self.select_file_ap_button = QPushButton("Select Excel File for AP")
        self.select_file_ap_button.setObjectName("ap_button")  # Setting object name for styling
        self.select_file_ap_button.clicked.connect(self.open_excel_file_ap)
        layout.addWidget(self.select_file_ap_button)

        self.select_file_worker_button = QPushButton("Select Excel File for Worker")
        self.select_file_worker_button.setObjectName("worker_button")  # Setting object name for styling
        self.select_file_worker_button.clicked.connect(self.open_excel_file_worker)
        layout.addWidget(self.select_file_worker_button)

        # Drop-down menu
        self.dropdown_menu = QComboBox()
        self.dropdown_menu.setObjectName("dropdown_menu")
        self.dropdown_menu.addItems([])
        layout.addWidget(self.dropdown_menu)

        # Text input box
        self.input_box = QLineEdit()
        self.input_box.setObjectName("input_box")
        self.input_box.setPlaceholderText("Enter the name of the file to be saved")
        layout.addWidget(self.input_box)

        self.run_button = QPushButton("Run Process")
        self.run_button.setObjectName("run_button")  # Setting object name for styling
        self.run_button.clicked.connect(self.run_button_process)
        layout.addWidget(self.run_button)

        # Apply CSS styles using Qt stylesheet
        self.setStyleSheet("""
                    QMainWindow {
                        background-color: #f0f0f0;
                        padding: 20px;
                    }
                    QLabel {
                        font-size: 14px;
                        color: #333;
                        margin-bottom: 10px;
                    }
                    QPushButton {
                        background-color: #008CBA;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        text-align: center;
                        text-decoration: none;
                        display: inline-block;
                        font-size: 14px;
                        margin: 5px;
                        cursor: pointer;
                        border-radius: 3px;
                        transition: background-color 0.3s ease;
                    }
                    QPushButton:hover {
                        background-color: #007B9A;
                    }
                    #worker_button {
                        background-color: #4CAF50;
                    }
                    #worker_button:hover {
                        background-color: #45a049;
                    }
                    #run_button {
                        background-color: #ff5722;
                    }
                    #run_button:hover {
                        background-color: #e64a19;
                    }
                    QComboBox, QLineEdit {
                        font-size: 14px;
                        padding: 10px;
                        margin-top: 10px;
                        border: 1px solid #ccc;
                        border-radius: 3px;
                    }
                """)

        self.selected_file_ap = None
        self.selected_file_worker = None
        self.entity = None
        self.output_name = None

    def open_excel_file_ap(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)",
                                                   options=options)
        if file_name:
            self.selected_file_ap = file_name
            self.ap_file_label.setText(f"Selected AP file: {file_name}")
            self.selected_file_worker = "Select Excel File"
            self.worker_file_label.setText(f"Selected Worker file: {self.selected_file_worker}")
            self.check_files_selected()

    def open_excel_file_worker(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)",
                                                   options=options)
        if file_name:
            self.selected_file_worker = file_name
            self.worker_file_label.setText(f"Selected Worker file: {file_name}")

    def check_files_selected(self):
        if self.selected_file_ap:
            df = input_file.get_file(self.selected_file_ap)
            self.dropdown_menu.clear()
            self.dropdown_menu.addItems(["select one Company/University/Hochschule"])
            lista = input_file.get_all_names(df)
            self.dropdown_menu.addItems(lista)
            self.input_box.clear()

    def capture_output_name(self):
        self.output_name = self.input_box.text()
        print(f"Captured input: {self.output_name}")
        filename = self.output_name + ".pdf"
        if os.path.exists(filename):
            os.remove(filename)

    def capture_entity(self):
        self.entity = self.dropdown_menu.currentText()

    def run_button_process(self):
        df = input_file.get_file(self.selected_file_ap)
        self.capture_output_name()
        self.capture_entity()
        if self.selected_file_ap is not None and self.selected_file_worker is not None and self.entity != "select one Company/University/Hochschule" and self.output_name != "":
            run_process(df, self.selected_file_ap, self.selected_file_worker, self.output_name, self.entity)
        else:
            show_popup(f"AP file ,Worker file, Output Name or Entity not selected")


def month_number_to_name(month_number):
    """
    Maps an integer representing a month number to the corresponding month name.

    Args:
        month_number (int): An integer from 1 to 12 representing the month.

    Returns:
        str: The name of the month.
    """
    month_names = [
        "Januar", "Februar", "Marz", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember"
    ]

    if 1 <= month_number <= 12:
        return month_names[month_number - 1]
    else:
        raise ValueError("Month number must be between 1 and 12.")


def run_process(df, filepath, filepath_workers, name_of_output_file, entity):
    # create instance of AP to access functions
    ap1 = AP.AP()

    # get path to file PM and Workers
    input_file.get_arbeitspaket(df)
    input_file.get_all_names(df)

    # get in which colum each pms starts
    lista = input_file.get_dates(filepath)

    # get when each pm starts and ends, how many months per working year and in which year each pm starts and ends
    list_datas, lista_months, list_begin, list_end = input_file.get_dates_unix(df, lista)

    # input_file.get_color_of_company(df,filepath,entity)

    lista_datas_not_to_change = list_datas

    # save dates
    ap1.add_dates(list_datas[0], list_datas[1])

    # get which PM id this company will have to do
    ap1.get_hours(input_file.get_Company(df, entity))

    # get all the id PM titles names and take the first ad last one out (may refactor this later)
    ap1.Nr = input_file.get_arbeitspaket(df)
    ap1.Nr = ap1.Nr[1:]
    # ap1.Nr = ap1.Nr[0:-1]

    # get ids of pm
    ids = input_file.get_nrs(df)

    # clear workers info
    worker.list_of_workers.clear()

    # get workers hours
    input_file.get_workers_info(filepath_workers, lista_months)

    # sorte the worker from most expensive to least
    worker.sorte_workers()

    repeted_wh_ids = []

    # Check if intervals span across multiple years
    ap1.check_if_same_years(ids, repeted_wh_ids, list_begin, list_end)

    # get start and end year and save on ap1.year_start, ap1.year_end
    ap1.get_smallest_year()
    ap1.get_biggest_year()

    # cleaning dict for zettel infos
    AP.global_data_zettel_infos.clear()

    hours_year_work_every_one = []
    worker_hours = []
    for wk in worker.list_of_workers:
        worker_hours = []
        for i in range(0, ap1.year_end - ap1.year_start + 1):
            worker_hours.append(float(wk.hours_available[i].item()))
        hours_year_work_every_one.append(worker_hours)

    pre_define_workers = input_file.get_workers_pre_defined(df)

    years = np.linspace(ap1.year_start, ap1.year_end, ap1.year_end - ap1.year_start + 1)
    list_aps = []
    dict_aps_infos = {}
    cost = 0

    for times in range(100):

        # clear workers info
        worker.list_of_workers.clear()

        # get workers hours
        input_file.get_workers_info(filepath_workers, lista_months)

        # sorte the worker from most expensive to least
        worker.sorte_workers()

        New_Nrs, New_ids, New_year_start, New_year_end, New_pre_define_workers, New_hours, Shuffle_to_Original_Index = AP.shuffle_aligned_lists(ap1.Nr, ids, lista_datas_not_to_change[0], lista_datas_not_to_change[1], pre_define_workers, ap1.hours)

        h, ids_check, Nrs, pre_def = ap1.get_workers([New_year_start,New_year_end], New_ids, ap1.year_start, ap1.year_end, New_Nrs,
                                                 entity,
                                                 df, New_pre_define_workers, New_hours)

    #h, ids_check, Nrs, pre_def = ap1.get_workers(lista_datas_not_to_change, ids, ap1.year_start, ap1.year_end, ap1.Nr,
    #                                             entity,
    #                                             df, pre_define_workers, ap1.hours)

        cost_this_version = calculate_proj_cost(years, hours_year_work_every_one)

        if cost_this_version > cost:
            cost = cost_this_version
            list_aps.clear()
            list_aps = AP.order_aps.copy()
            dict_aps_infos = copy.deepcopy(AP.global_data_zettel_infos)
        AP.order_aps.clear()
        AP.global_data_zettel_infos.clear()



    # For example, reconstruct data arrays
    AP.order_aps = list_aps
    AP.global_data_zettel_infos = dict_aps_infos
    order_aps_final = sorted(AP.order_aps, key=lambda x: tuple(map(float, x[1].split("."))))

    restored_Nrs = [x[0] for x in order_aps_final]
    restored_ids = [x[1] for x in order_aps_final]
    restored_start = [x[2] for x in order_aps_final]
    restored_end = [x[3] for x in order_aps_final]
    restored_dates = [x[4] for x in order_aps_final]
    restored_hours = [x[5] for x in order_aps_final]
    restored_workers = [x[6] for x in order_aps_final]




    html_content_1 = """
    <html lang="de">
    <head>
        <meta charset="UTF-8">
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 20px;
            }
            h1 {
                color: #333;
                text-align: center;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 20px;
                table-layout: fixed;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
                overflow: hidden;
                white-space: normal;
                word-break: normal;      /* Only break at natural points */
                hyphens: auto;           /* Allow hyphenation */
            }
            th {
                background-color: #f2f2f2;
                color: #333;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            tr:hover {
                background-color: #f5f5f5;
            }
            td:nth-child(2) {
                font-size: 12px;  /* Adjust the font size as needed */
            }
        </style>
    </head>
    <body>
        <h1>Arbeitspaketbericht</h1>
        <table>
            <colgroup>
                <col style="width:10%">    <!-- Id -->
                <col style="width:37%">   <!-- AP (WIDER) -->
                <col style="width:17%">   <!-- Startdatum -->
                <col style="width:17%">   <!-- Enddatum -->
                <col style="width:10%">    <!-- Id Arbeiter -->
                <col style="width:8%">    <!-- WH -->
            </colgroup>
            <tr>
                <th>Id</th>
                <th>AP</th>
                <th>Startdatum</th>
                <th>Enddatum</th>
                <th>Id Arbeiter</th>
                <th>WH</th>
            </tr>
    """

    sum_test = 0
    sum_test2 = 0
    ap_not_distribute = []
    array_working_hours_per_year = np.zeros((len(worker.list_of_workers), len(years)))

    for w, wh, dates_st, dates_ft, Nr, id in zip(restored_workers,restored_hours,restored_start,restored_end,
                                                 restored_Nrs,restored_ids):
        html_content_1 += f"""
                                <tr style="background-color: {"#ccffcc"};">
                                    <td>{id}</td>
                                    <td>{Nr}</td>
                                    <td>{dates_st}</td>
                                    <td>{dates_ft}</td>
                                    <td>{w.id}</td>
                                    <td>{round(wh,2)}</td>
                                </tr>
                        """
    html_content_1 += """
                    </table>
                    <div style="page-break-before: always;"></div>
                </body>
                """

    # Generate HTML content with styling for the second table
    html_content_1 += """
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, sans-serif;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                th, td {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 10px;
                    font-size: 12px;
                }
                th {
                    background-color: #f2f2f2;
                }
            </style>
        </head>
        <body>
            <h1>Summen arbeiterbericht</h1>
            <table>
                <tr>
                    <th>Jahr</th>
        """

    for i in range(len(worker.list_of_workers)):
        html_content_1 += f"""
                        <th>Summen arbeiter {i + 1}</th>
            """
    html_content_1 += f"""
                </tr>
        """

    p_p_y = np.array(lista_months) / 12

    sum_t = 0
    cost_project = 0

    new_array = []
    new_array_hours = []
    while len(worker.list_of_workers) != 0:
        lowest_index_elem = worker.Worker(1000, 0, 0, 0, "", "",0)
        for element in worker.list_of_workers:
            if element.id < lowest_index_elem.id:
                lowest_index_elem = element
        new_array.append(lowest_index_elem)
        index = worker.list_of_workers.index(lowest_index_elem)
        new_array_hours.append(hours_year_work_every_one[index])

        worker.list_of_workers.remove(lowest_index_elem)
        hours_year_work_every_one.remove(hours_year_work_every_one[index])

    worker.list_of_workers = new_array
    hours_year_work_every_one = new_array_hours

    # Add rows for sum worker data
    for i in range(len(years)):
        html_content_1 += f"<tr>"
        html_content_1 += f"<td>{int(years[i])}</td>"
        for j in range(len(worker.list_of_workers)):
            html_content_1 += f"<td>{round((hours_year_work_every_one[j][i]) - (worker.list_of_workers[j].hours_available[i][0]),2)}</td>"
            sum_t += hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]
            cost_project += round(float(hours_year_work_every_one[j][i] - worker.list_of_workers[j].hours_available[i][0]) * \
                            worker.list_of_workers[j].salary,2)
        html_content_1 += f"</tr>"

    # Add a row for total hours
    html_content_1 += "<tr>"
    html_content_1 += "<td><strong>Total</strong></td>"

    workers_total_hours = []
    index_year = 0

    for workers_t in worker.list_of_workers:
        w_hours = []
        for index_ele in range(len(workers_t.hours_available)):
            w_hours.append(workers_t.hours_available[index_ele][0])
        workers_total_hours.append(w_hours)

    total_w = []
    for index_ele in range(len(workers_total_hours)):
        total_w.append(sum(np.array(hours_year_work_every_one[index_ele]) - np.array(workers_total_hours[index_ele])))



    # Add totals for each worker
    for total in total_w:
        html_content_1 += f"<td><strong>{round(total,2)}</strong></td>"

    # Close the table and HTML body for the second table
    html_content_1 += """
            </table>
            <table>
                <tr>
                   <th>Summe der Gesamtstunden</th>
                   <th>Stunden nicht verteilt</th>
                   <th>APs nicht verteilt</th>
                   <th>Projektkosten</th>
                   <th>Anzahl der APs</th>
                </tr>
        """

    sum_t_b = sum_t
    cost_project_formatted = format_euros(cost_project)
    aps_str = ""

    for aps in ap_not_distribute:
        aps_str += aps
        aps_str += ", "

    aps_str = aps_str[:-2]

    if aps_str == "":
        aps_str = "Alle APs verteilt"

    html_content_1 += f"""
            <tr>
                <td>{round(sum_t_b,2)}</td>
                <td>{sum_test}</td>
                <td>{aps_str}</td>
                <td>{cost_project_formatted}</td>
                <td>{len(h)}</td>
            </tr>
        """

    html_content_1 += """
        </table>
        </body>
        </html>
        """

    # Generate HTML content with styling for the second table
    html_content_2 = ""
    html_content_2 += """
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body {
                        font-family: Arial, sans-serif;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                    }
                    th, td {
                        border: 1px solid #dddddd;
                        text-align: left;
                        padding: 10px;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                </style>
            </head>
            <body>
                <h1>Terminverteilung</h1>
                <table>
                    <tr>
                        <th>Arbeiter</th>
                        <th>AP-Id</th>
                        <th>Monat</th>
                        <th>Jahr</th>
                        <th>Stunden</th>
                        <th>PM (wird zur Stundenberechnung verwendet: 1 PM = 160 Stunden) </th>
                    </tr>
            """

    #sorted_entries = sorted(
    #    AP.global_data_zettel_infos.items(),  # Sort by worker_id
    #    key=lambda x: (
    #        x[0],  # Sort by worker_id (x[0] is the worker_id)
    #        [entry['AP id'] for entry in x[1]],
    #        [entry['year'] for entry in x[1]],  # Sort by year (x[1] contains the entries for each worker)
    #        [entry['month'] for entry in x[1]]  # Sort by month (x[1] contains the entries for each worker)
    #    )
    #)

    all_entries = [
        entry
        for worker_entries in AP.global_data_zettel_infos.values()
        for entry in worker_entries
    ]

    sorted_entries = sorted(
        all_entries,
        key=lambda e: (
            e['worker_id'],
            ap_id_sort_key(e['AP id']),  # usa parse para ordenar, mas não altera o valor
            month_map.get(e['month'], 0),  # mês como número
            e['year'],
            e['hours']
        )
    )




    for entry in sorted_entries:
        # Extract month, hours, and PM
        month = entry['month']
        hours = entry['hours']
        year = entry['year']
        AP_id = entry['AP id']
        worker_id = entry['worker_id']

        if hours == 0:
            continue

        name = ""

        for wks in worker.list_of_workers:
            if wks.id == worker_id:
                name = str(wks.name) + " " + str(wks.surname)

        # Add a row for each entry
        html_content_2 += f"""
         <tr>
            <td>{name}</td>
            <td>{AP_id}</td>
            <td>{get_german_month(month)}</td>
            <td>{year}</td>
            <td>{round(hours,2) * 160}</td>
            <td>{round(hours,2)}</td>
        </tr>
        """
    html_content_2 += """
             </table>
         </body>
         </html>
         """

    # Generate HTML content with styling for the second table
    html_content_2 += """
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, sans-serif;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                th, td {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 10px;
                }
                th {
                    background-color: #f2f2f2;
                }
            </style>
        </head>
        <body>
            <h1>Monatlicher Arbeiterbericht</h1>
            <table>
                <tr>
                    <th>Arbeiter</th>
                    <th>Stunden</th>
                    <th>Jahr</th>
                    <th>Monat</th>
                 </tr>
        """

    months_german = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober",
                     "November", "Dezember"]

    last_month = 12
    for wk in worker.list_of_workers:
        first_year = 0
        for year_idx, year in enumerate(years):
            if first_year == 0:
                months_to_iterate = 12 - lista_months[0]
                first_year = 1
                last_month = 12
            else:
                months_to_iterate = 0
                last_month = lista_months[year_idx]

            for i in range(months_to_iterate, last_month):
                month_idx = i
                hours = 1 - wk.hours_available_per_month[year_idx][month_idx]

                html_content_2 += f"""
                <tr>
                    <td>{str(wk.name) + " " + str(wk.surname)}</td>
                    <td>{round(hours * 40,2)}</td>
                    <td>{int(year)}</td>
                    <td>{months_german[month_idx]}</td>
                </tr>
                """

    html_content_2 += """
            </table>
        </body>
        </html>
    """

    # Save HTML content to a file
    with open("output.html", "w") as file:
        file.write(html_content_1)

    with open("output2.html", "w") as file:
        file.write(html_content_2)

    if len(name_of_output_file) == 0:
        print("Error name of pdf, it cannot be empty")
        exit(1)

    if len(name_of_output_file) > 100:
        print("Error name of pdf, it is way too big")
        exit(1)

    pdf_output = BytesIO()
    pdf_output_2 = BytesIO()

    HTML('output.html').write_pdf(name_of_output_file + "_" + entity + "_datum" + ".pdf")
    HTML('output2.html').write_pdf(name_of_output_file + "_" + entity + "_organizer"+".pdf")
    print("acabou")



def round_down_to_step(value, step):
    return round((int(value / step)) * step, 10)

def round_up_to_step(value, step):
    if value % step == 0:
        return value
    return ((int(value / step)) + 1) * step

def round_to_step(value, step):
    """
    Rounds value up to the nearest multiple of `step`,
    but first rounds down to nearest 0.05.
    """
    value = round_down_to_step(value, 0.05)  # optional pre-rounding
    return round_up_to_step(value, step)

def value_to_color(value):
    """Return a color based on the value. Near 0 is red, in the middle is blue, near 1 is green."""
    # Ensure value is between 0 and 1
    value = max(0, min(1, value))

    if value < 0.5:
        red = int((1 - value * 2) * 255)
        blue = int(value * 2 * 255)
        return f"rgb({red}, 0, {blue})"
    else:
        blue = int((1 - value) * 2 * 255)
        green = int((value - 0.5) * 2 * 255)
        return f"rgb(0, {green}, {blue})"


def format_euros(amount):
    return '€{:,.2f}'.format(amount)


def allocate_value(array, start_date, end_date, worker_id, value, years):
    # Convert start and end dates to datetime objects
    start_date = datetime.strptime(start_date, "%d.%m.%Y")
    end_date = datetime.strptime(end_date, "%d.%m.%Y")

    # Get the total number of days between start and end dates
    total_days = (end_date - start_date).days + 1

    # Loop over each year and calculate the value allocation
    for year in years:
        # Calculate the start and end dates of the current year
        year_int = int(year)
        year_start = datetime(year_int, 1, 1)
        year_end = datetime(year_int, 12, 31)

        # Determine the overlap of the year with the given start and end dates
        overlap_start = max(start_date, year_start)
        overlap_end = min(end_date, year_end)

        # Calculate the number of overlapping days in the year
        overlapping_days = (overlap_end - overlap_start).days + 1
        if overlapping_days > 0:
            # Calculate the portion of the value for this year
            year_allocation = (overlapping_days / total_days) * value
            # Assign the calculated value to the array
            year_index = int(int(year) - years[0])
            array[worker_id - 1][year_index] += year_allocation

    return array

# Function to open the PDF after it's created
def open_pdf(file_path):
    if os.name == 'posix':  # For Linux and macOS
        os.system(f'open "{file_path}"')
    elif os.name == 'nt':  # For Windows
        os.system(f'start {file_path}')
    else:
        print("Unsupported OS")


def main():
    # get reference for period file
    # df = input_file.get_file(filepath)

    app = QApplication(sys.argv)
    excel_reader_app = ExcelReaderApp()
    excel_reader_app.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
