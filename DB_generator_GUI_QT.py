# -*- coding: utf-8 -*-
"""
Created on Thu Feb 15 11:12:55 2024

@author: boris.sorokin@skao.int

This is a next version of DB generator GUI with QT interface. Since it is supposed to be under GPL, should be fine
"""


import sys
import pyodbc
import os
import csv
import docx
from docx.enum.section import WD_ORIENT

from PyQt5.QtWebEngineWidgets import QWebEngineView

from PyQt5.QtWidgets import (QApplication, QMainWindow, QDesktopWidget, QWidget, QPushButton, QFileDialog, QLabel,
                             QMessageBox, QGridLayout, QGroupBox, QDialog, QTableWidget, QTableWidgetItem, QCheckBox,
                             QHBoxLayout, QProgressDialog)

from PyQt5.QtGui import QIcon, QPixmap, QDesktopServices

from PyQt5.QtCore import Qt, QParallelAnimationGroup, QPropertyAnimation, QRect, QEventLoop, QEasingCurve, QUrl

class MainApp(QMainWindow):
    # This is the main window with database selector, interactive database, and saving capability
    def __init__(self):
        super().__init__()
        self.dbConnection = None
        self.interactive_database = None
        self.desired_width = 1280
        self.desired_height = 720
        self.initUI()

    def initUI(self):
        # Set the window properties
        self.setWindowTitle('IAU CPS RAS Database Tool')
        self.setWindowIcon(QIcon('cps-logo-mono.ico'))

        # Center the window on the screen or parent
        self.centerWindow(self)
        self.animateOpening(self, self.desired_width, self.desired_height)

        # Create menubar
        menuBar = self.menuBar()
        helpMenu = menuBar.addMenu('Help')

        aboutAction = helpMenu.addAction('About')
        aboutAction.triggered.connect(self.show_about)

        # Create a group box for ITU Database tools
        self.ituToolsGroup = QGroupBox('ITU Database Tools')
        gridLayoutITU = QGridLayout()

        # Create a button for selecting the database file
        self.button_open = QPushButton('Select ITU database to import', self)
        self.button_open.clicked.connect(self.database_select)
        # Place the button in the grid layout
        gridLayoutITU.addWidget(self.button_open, 0, 0)

        # Create a status indicator for selecting the file
        self.statusLight_open = QLabel(self)
        self.statusLight_open.setFixedSize(20, 20)
        self.updateStatusLight(self.statusLight_open,
                               False, 'Database not selected.')
        gridLayoutITU.addWidget(self.statusLight_open, 0, 1)
        gridLayoutITU.setColumnStretch(1, 0)

        # Create a button for connecting to the selected database
        self.button_connect = QPushButton('Connect to selected database', self)
        self.button_connect.clicked.connect(self.database_connect)
        # Place the button in the grid layout
        gridLayoutITU.addWidget(self.button_connect, 1, 0)
        self.button_connect.setEnabled(False)
        self.button_connect.setToolTip('Select a database first.')

        # Create a status indicator for selecting the file
        self.statusLight_connect = QLabel(self)
        self.statusLight_connect.setFixedSize(20, 20)
        self.updateStatusLight(self.statusLight_connect,
                               False, 'Database not connected.')
        gridLayoutITU.addWidget(self.statusLight_connect, 1, 1)
        gridLayoutITU.setColumnStretch(1, 0)

        # Create a button for showing the interactive list
        self.button_show_list = QPushButton(
            'Show the list of radio astronomy stations', self)
        self.button_show_list.clicked.connect(self.interactive_database_show)
        # Place the button in the grid layout
        gridLayoutITU.addWidget(self.button_show_list, 2, 0)
        self.button_show_list.setEnabled(False)
        self.button_show_list.setToolTip('Connect a database first.')

        self.ituToolsGroup.setLayout(gridLayoutITU)

        # Create a group box for ITU Database tools
        self.exportToolsGroup = QGroupBox('Export Tools')
        gridLayoutExport = QGridLayout()

        self.button_export_csv = QPushButton(
            'Export all data as CSV', self)
        self.button_export_csv.clicked.connect(self.save_csv)
        self.button_export_csv.setEnabled(False)
        self.button_export_csv.setToolTip('Connect a database first.')
        gridLayoutExport.addWidget(self.button_export_csv, 0, 0)

        self.button_export_word = QPushButton(
            'Export all data as DOCX', self)
        self.button_export_word.clicked.connect(self.save_word)
        self.button_export_word.setEnabled(False)
        self.button_export_word.setToolTip('Connect a database first.')
        gridLayoutExport.addWidget(self.button_export_word, 0, 1)

        self.exportToolsGroup.setLayout(gridLayoutExport)

        # Setting central widget
        centralWidget = QWidget()
        self.setCentralWidget(centralWidget)

        # Main layout for the widget
        mainLayout = QGridLayout(centralWidget)
        mainLayout.addWidget(self.ituToolsGroup, 0, 0)
        mainLayout.addWidget(self.exportToolsGroup, 1, 0)

        self.statusBar().showMessage('    Ready')
        self.statusBar().setStyleSheet("""
            QStatusBar {
                background-color: #404040;
                color: white;
                border: 2px solid black;
            }
        """)
        self.show()
        self.setFocus()
        self.raise_()
        self.activateWindow()

    def database_select(self):
        # Selecting database callback
        options = QFileDialog.Options()
        self.database_file_name, _ = QFileDialog.getOpenFileName(
            self, "QFileDialog.getOpenFileName()", "", "MDB Files (*.mdb)", options=options)
        if self.database_file_name:
            self.statusBar().showMessage(
                f'    Database file selected: {self.database_file_name}')
            self.updateStatusLight(self.statusLight_open,
                                   True, 'Database selected')
            self.button_connect.setEnabled(True)
            self.button_connect.setToolTip(None)
        else:
            self.updateStatusLight(self.statusLight_open,
                                   False, 'Database not selected')
            self.button_connect.setEnabled(False)
            self.button_connect.setToolTip('Select a database first.')

            self.button_export_csv.setEnabled(False)
            self.button_export_csv.setToolTip('Select a database first.')

            self.button_export_word.setEnabled(False)
            self.button_export_word.setToolTip('Select a database first.')

    def database_connect(self):
        # Attempt to connect to the selected database
        try:
            connectionString = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.database_file_name}'
            try:
                self.dbConnection.close()
            except:
                pass

            self.dbConnection = pyodbc.connect(connectionString)
            self.updateStatusLight(
                self.statusLight_connect, True, 'Database connected')
            self.statusBar().showMessage('    Database connected. Checking version...')
            SQL = "SELECT d_create, comment FROM srs_ooak"
            rows = self.parse_database(SQL)[0]
            self.database_date = rows[0]
            self.database_version = rows[1][0:7]
            self.statusBar().showMessage(
                f'Connected to database {self.database_version} published on {self.database_date.date()}')
            self.button_show_list.setEnabled(True)
            self.button_show_list.setToolTip(None)

            self.button_export_csv.setEnabled(True)
            self.button_export_csv.setToolTip(None)

            self.button_export_word.setEnabled(True)
            self.button_export_word.setToolTip(None)
        except Exception as e:
            QMessageBox.critical(self, "Database Connection Error",
                                 f"An error occurred while connecting to the database:\n{e}")
            self.updateStatusLight(
                self.statusLight_connect, True, 'Database connection error')
            self.statusBar().showMessage('    Database connection error')
            self.button_show_list.setToolTip('Connect a database first.')

    def interactive_database_show(self):
        self.animateClosing(self)
        self.showMinimized()
        self.setEnabled(False)
        self.interactive_database = InteractiveDatabase(self)

    def parse_database(self, SQL):
        """Parse the database and display results in a new window and save to a Word document."""
        rows = []
        if self.dbConnection:
            try:
                cursor = self.dbConnection.cursor()
                cursor.execute(SQL)
                rows = cursor.fetchall()
                cursor.close()
            except Exception as e:
                QMessageBox.critical(self,
                                     "Database Error", f"Error parsing database: {e}")
        return rows

    def show_about(self):
        # Show about window.
        aboutDialog = AboutDialog(self)
        aboutDialog.exec_()

    def save_csv(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save as CSV", f"RAS_DB_FULL_CSV_{self.database_version}_{self.database_date.date()}", "CSV Files (*.csv)")
        if filePath:
            if os.path.basename(filePath) == 'geographical-areas.csv':
                QMessageBox.critical(
                    self, "Error", "This file is an important app file and cannot be overwritten.")
                return

            SQL = 'SELECT ntc_id, adm, ctry, stn_name, long_dec, lat_dec FROM com_el WHERE ntc_type=\'R\' ORDER BY adm asc, stn_name asc;'
            com_el_rows = self.parse_database(SQL)
            station_number = len(com_el_rows)

            fields = ['Notice ID', 'Administration', 'Region/Location', 'Station name',
                      'Longitude', 'Latitude', 'Longitude Degrees',
                      'Longitude East/West', 'Longitude minutes', 'Longitude seconds',
                      'Latitude Degrees', 'Latitude North/South', 'Latitude minutes',
                      'Latitude seconds', 'Elevation minimum', 'Elevation maximum',
                      'Azimuth from', 'Azimuth to', 'Beam name', 'Antenna pattern ID',
                      'Antenna pattern Name', 'Centre frequency, MHz',
                      'Group ID', 'Noise Temp, K', 'Frequency minimum, MHz',
                      'Frequency maximum, MHz', 'Date brought into use', 'Date received',
                      'IFIC no (wic_no)', 'Date updated']

            with open(filePath, 'w', newline='') as file:
                csv_writer = csv.writer(file, delimiter=',')
                csv_writer.writerow(fields)
                for index in range(0, station_number):
                    ntc_id = com_el_rows[index][0]
                    SQL = 'SELECT long_deg, long_ew, long_min, long_sec, lat_deg, lat_ns, lat_min, lat_sec, elev_min, elev_max, azm_fr, azm_to FROM e_stn WHERE ntc_id=' + \
                        str(ntc_id)+';'
                    e_stn_rows = self.parse_database(SQL)
                    SQL = 'SELECT beam_name, pattern_id FROM e_ant WHERE ntc_id=' + \
                        str(ntc_id)+';'
                    e_ant_rows = self.parse_database(SQL)

                    beam_number = len(e_ant_rows)
                    beam_names = []
                    ant_id = []
                    ant_names = []

                    for subindex_beam in range(0, beam_number):
                        beam_names.append(e_ant_rows[subindex_beam][0])
                        ant_id.append(e_ant_rows[subindex_beam][1])
                        if ant_id[-1] == None:
                            SQL = 'SELECT attch_e, ant_diam FROM e_ant WHERE ntc_id=' + \
                                str(ntc_id)+';'
                            local_ant_rows = self.parse_database(SQL)
                            if local_ant_rows[subindex_beam][1] == None:
                                ant_names.append('NonTypical, see attachment {} to the relevant IFIC for details.'.format(
                                    local_ant_rows[subindex_beam][0]))
                            else:
                                ant_names.append('NonTypical, submitted diameter is {} meters, see attachment {} to the relevant IFIC for details.'.format(
                                    local_ant_rows[subindex_beam][1], local_ant_rows[subindex_beam][0]))
                        else:
                            SQL = 'SELECT pattern FROM ant_type WHERE pattern_id=' + \
                                str(ant_id[-1])+';'
                            ant_type_rows = self.parse_database(SQL)
                            try:
                                ant_names.append(ant_type_rows[0][0])
                            except:
                                SQL = 'SELECT attch_e, ant_diam FROM e_ant WHERE ntc_id=' + \
                                    str(ntc_id)+';'
                                local_ant_rows = self.parse_database(SQL)
                                if local_ant_rows[subindex_beam][1] == None:
                                    ant_names.append('NonTypical, see attachment {} to the relevant IFIC for details.'.format(
                                        local_ant_rows[subindex_beam][0]))
                                else:
                                    ant_names.append('NonTypical, submitted diameter is {} meters, see attachment {} to the relevant IFIC for details.'.format(
                                        local_ant_rows[subindex_beam][1], local_ant_rows[subindex_beam][0]))

                        SQL = 'SELECT grp_id, noise_t, freq_min, freq_max, d_inuse, d_rcv, wic_no, d_upd FROM grp WHERE (ntc_id='+str(
                            ntc_id)+'and beam_name=\''+beam_names[subindex_beam]+'\''+');'
                        grp_rows = self.parse_database(SQL)

                        SQL = 'SELECT freq_mhz FROM freq WHERE (ntc_id='+str(
                            ntc_id)+'and beam_name=\''+beam_names[subindex_beam]+'\''+');'
                        freq_rows = self.parse_database(SQL)

                        group_number = len(grp_rows)

                        for subindex_group in range(0, group_number):
                            if subindex_beam == 0 and subindex_group == 0:
                                next_line = list(com_el_rows[index])+list(e_stn_rows[0])+list([beam_names[subindex_beam]])+list([ant_id[subindex_beam]])+list(
                                    [ant_names[subindex_beam]])+list([str(freq_rows[subindex_group][0])])+list(grp_rows[subindex_group])
                                csv_writer.writerow(next_line)
                            else:
                                next_line = list([''])*18+list([beam_names[subindex_beam]])+list([ant_id[subindex_beam]])+list(
                                    [ant_names[subindex_beam]])+list([str(freq_rows[subindex_group][0])])+list(grp_rows[subindex_group])
                                csv_writer.writerow(next_line)

    def save_word(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save as DOCX", f"RAS_DB_FULL_DOCX_{self.database_version}_{self.database_date.date()}", "Word Files (*.docx)")
        if filePath:
            doc = docx.Document()
            section = doc.sections[0]
            section = doc.sections[0]
            section.orientation = WD_ORIENT.PORTRAIT
            doc.add_heading(
                "Annex 1. The list of radio astronomy stations known to the IAU CPS", level=1)
            doc.add_paragraph("This list is based on ITU-R IFIC database.")

            SQL = 'SELECT ntc_id, adm, ctry, stn_name, long_dec, lat_dec FROM com_el WHERE ntc_type=\'R\' ORDER BY adm asc, stn_name asc;'
            com_el_rows = self.parse_database(SQL)
            station_number = len(com_el_rows)

            progressDialog = QProgressDialog(
                "Operation in progress...", "Cancel", 0, station_number, self)
            progressDialog.setWindowTitle("Saving...")
            progressDialog.setWindowModality(Qt.WindowModal)
            progressDialog.setCancelButton(None)
            progressDialog.show()

            # Establish a link between ITU country codes and country names
            # https://www.itu.int/en/ITU-R/terrestrial/fmd/Pages/geo_area_list.aspx
            country_codes_to_names = {}
            with open('geographical-areas.csv', newline='') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    country_codes_to_names[row[0]] = row[1]

            for index, row in enumerate(com_el_rows):
                progressDialog.setLabelText(
                    f"Now populating station {index+1} of {station_number}")
                ntc_id = row[0]
                try:
                    admin_name = country_codes_to_names[row[1]]
                except:
                    admin_name = 'Unknown'
                try:
                    country_name = country_codes_to_names[row[2]]
                except:
                    country_name = 'Unknown'
                station_name = row[3]
                station_longitude = com_el_rows[index][4]
                station_latitude = com_el_rows[index][5]

                SQL = f'SELECT elev_min, ant_alt FROM e_stn WHERE ntc_id={ntc_id};'
                e_stn_rows = self.parse_database(SQL)

                station_min_elevation = e_stn_rows[0][0]
                if station_min_elevation == None:
                    station_min_elevation = 'N/A'
                    
                station_antenna_altitude = e_stn_rows[0][1]
                if station_antenna_altitude == None:
                    station_antenna_altitude = 'N/A'

                SQL = 'SELECT beam_name, ant_diam, gain FROM e_ant WHERE ntc_id=' + \
                    str(ntc_id)+';'
                e_ant_rows = self.parse_database(SQL)

                beam_number = len(e_ant_rows)

                beam_names = []
                noise_temp = []
                ant_diameter = []
                ant_gain = []
                freq_min = []
                freq_max = []

                for subindex_beam in range(0, beam_number):
                    beam_names.append(e_ant_rows[subindex_beam][0])
                    if e_ant_rows[subindex_beam][1] == None:
                        ant_diameter.append('N/A')
                    else:
                        ant_diameter.append(e_ant_rows[subindex_beam][1])

                    if e_ant_rows[subindex_beam][2] == None:
                        ant_gain.append('N/A')
                    else:
                        ant_gain.append(e_ant_rows[subindex_beam][2])

                    SQL = f"SELECT noise_t, freq_min, freq_max FROM grp WHERE (ntc_id={ntc_id} and beam_name='{beam_names[subindex_beam]}');"
                    grp_rows = self.parse_database(SQL)

                    noise_temp.append(grp_rows[0][0])
                    freq_min.append(grp_rows[0][1])
                    freq_max.append(grp_rows[0][2])

                station_freq_min = min(freq_min)
                station_freq_max = max(freq_max)

                # Flushing all gathered data to the document
                doc.add_heading(f'Station "{station_name}"', level=2)
                doc.add_heading('Overview', level=3)
                doc.add_paragraph(f'Station number: {index+1}')
                doc.add_paragraph(
                    f'Responsible administration: "{admin_name}"')
                doc.add_paragraph(f'Country/region location: "{country_name}"')
                doc.add_paragraph(f'Station short name: "{station_name}')
                doc.add_paragraph('Station long: name "N/A"')
                doc.add_paragraph('Station type: "N/A"')
                doc.add_paragraph(
                    f'Station longitude [deg]: "{station_longitude}"')
                doc.add_paragraph(
                    f'Station latitude [deg]: "{station_latitude}"')
                doc.add_paragraph(f'Station altitude (AMSL) "{station_antenna_altitude}"')
                doc.add_paragraph(
                    f'Minimum elevation [deg]: "{station_min_elevation}"')
                doc.add_paragraph('Operational "N/A"')
                doc.add_paragraph('Used for science "N/A"')
                doc.add_paragraph(
                    f'Minimum Station Frequency [MHz]: "{station_freq_min} MHz"')
                doc.add_paragraph(
                    f'Maximum Station Frequency [MHz]: "{station_freq_max} MHz"')
                doc.add_paragraph('Contact (website) "N/A"')
                doc.add_paragraph('Contact (address) "N/A"')
                doc.add_paragraph('Contact (phone) "N/A"')
                doc.add_paragraph('Contact (e-mail) "N/A"')

                doc.add_heading('Antenna information', level=3)

                for beam_index in range(0, beam_number):
                    doc.add_heading(f'Antenna #{beam_index+1}', level=4)
                    doc.add_paragraph('Feed/Rx height above ground [m] "N/A"')
                    doc.add_paragraph(
                        f'Noise temparature [K]: "{noise_temp[beam_index]}"')
                    doc.add_paragraph(
                        f'Antenna diameter [m]: "{ant_diameter[beam_index]}"')
                    doc.add_paragraph(
                        f'Maximum antenna gain [dBi]: "{ant_gain[beam_index]}"')
                    doc.add_paragraph(
                        f'Minimim antenna frequency [MHz]: "{freq_min[beam_index]}"')
                    doc.add_paragraph(
                        f'Maximum antenna frequency [MHz]: "{freq_max[beam_index]}"')
                    doc.add_paragraph('Cryocooled: "N/A"')
                    doc.add_paragraph('Supports RAS mode continuum: "N/A"')
                    doc.add_paragraph('Supports RAS mode spectroscopy: "N/A"')
                    doc.add_paragraph('Supports RAS mode VLBI: "N/A"')

                doc.add_page_break()
                progressDialog.setValue(index+1)

            try:
                doc.save(filePath)
            except Exception as e:
                QMessageBox.critical(self, "Docx saving Error",
                                     f"An error occurred while preparing the docx file:\n{e}")

    def updateStatusLight(self, widget, status, status_text):
        # Update the status light color based on whether connection was successfully established
        if status:
            widget.setStyleSheet(
                "QLabel { background-color: green; border-radius: 10px; }")
        else:
            widget.setStyleSheet(
                "QLabel { background-color: red; border-radius: 10px; }")

        self.connectionStatus = status_text
        widget.setToolTip(status_text)

    def centerWindow(self, window, parent=None):
        # Center the window on the screen or parent window
        if parent:
            parentGeometry = parent.frameGeometry()
            centerPoint = parentGeometry.center()
            window.move(centerPoint - window.rect().center())
        else:
            screenGeometry = QDesktopWidget().screenGeometry()
            centerPoint = screenGeometry.center()
            window.move(centerPoint - window.rect().center())

    def animateOpening(self, window, desired_width, desired_height, parent=None):
        # This function animates opening of the window

        # Block interaction
        window.setEnabled(False)

        if parent:
            centerPoint = parent.frameGeometry().center()
        else:
            desktopWidget = QDesktopWidget()
            centerPoint = desktopWidget.availableGeometry(
                desktopWidget.primaryScreen()).center()

        startRect = QRect(0, 0, int(desired_width * 0.1),
                          int(desired_height * 0.1))
        startRect.moveCenter(centerPoint)
        endRect = QRect(0, 0, desired_width, desired_height)
        endRect.moveCenter(centerPoint)

        # Create and configure the geometry animation
        window.animation_geometry_opening = QPropertyAnimation(
            window, b"geometry")
        window.animation_geometry_opening.setDuration(
            600)  # Duration of animation in ms
        window.animation_geometry_opening.setStartValue(startRect)
        window.animation_geometry_opening.setEndValue(endRect)
        window.animation_geometry_opening.setEasingCurve(
            QEasingCurve.InOutCubic)

        # Create and configure the opacity animation
        window.animation_opacity_opening = QPropertyAnimation(
            window, b"windowOpacity")
        window.animation_opacity_opening.setDuration(500)
        window.animation_opacity_opening.setStartValue(-0.5)
        window.animation_opacity_opening.setEndValue(1.0)
        window.animation_opacity_opening.setEasingCurve(
            QEasingCurve.InOutCubic)

        # Grouping animations
        window.animation_group_opening = QParallelAnimationGroup()
        window.animation_group_opening.addAnimation(
            window.animation_geometry_opening)
        window.animation_group_opening.addAnimation(
            window.animation_opacity_opening)

        window.animation_geometry_opening.finished.connect(
            lambda: window.setEnabled(True))
        window.animation_group_opening.start()

    def animateClosing(self, window):
        # This function animates closing of the window.

        # Block interactions
        window.setEnabled(False)

        # Get current size and position
        currentRect = window.geometry()
        centerPoint = currentRect.center()

        # Calculate the final rectangle (10% of the current size)
        endRect = QRect(0, 0, int(currentRect.width() * 0.1),
                        int(currentRect.height() * 0.1))
        endRect.moveCenter(centerPoint)

        # Create and configure the geometry animation
        window.animation_geometry_closing = QPropertyAnimation(
            window, b"geometry")
        window.animation_geometry_closing.setDuration(
            500)  # Duration in milliseconds
        window.animation_geometry_closing.setStartValue(currentRect)
        window.animation_geometry_closing.setEndValue(endRect)
        window.animation_geometry_closing.setEasingCurve(
            QEasingCurve.InOutCubic)

        # Create and configure the opacity animation
        window.animation_opacity_closing = QPropertyAnimation(
            window, b"windowOpacity")
        window.animation_opacity_closing.setDuration(500)
        window.animation_opacity_closing.setStartValue(1.0)
        window.animation_opacity_closing.setEndValue(-0.5)
        window.animation_opacity_closing.setEasingCurve(
            QEasingCurve.InOutCubic)

        # Grouping animations
        window.animation_group_closing = QParallelAnimationGroup()
        window.animation_group_closing.addAnimation(
            window.animation_geometry_closing)
        window.animation_group_closing.addAnimation(
            window.animation_opacity_closing)

        loop = QEventLoop()
        window.animation_geometry_closing.finished.connect(window.hide)
        window.animation_geometry_closing.finished.connect(loop.quit)
        window.animation_group_closing.start()
        loop.exec_()

    def closeEvent(self, event):
        # Check if the database connection is active and close it before exiting
        self.animateClosing(self)

        if self.dbConnection:
            try:
                self.dbConnection.close()
                self.statusBar().showMessage('    Database connection closed. App will soon close')
            except Exception:
                self.statusBar().showMessage(
                    '    Database connection was already closed. App will soon close')
        event.accept()


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent

        self.initUI()

    def initUI(self):
        self.setWindowTitle("About")
        self.setWindowIcon(QIcon('cps-logo-mono.ico'))
        self.setWindowFlags(self.windowFlags(
        ) & ~Qt.WindowContextHelpButtonHint | Qt.WindowCloseButtonHint)
        self.parent.animateOpening(self, desired_width=640, desired_height=360)
        self.setModal(True)

        layout = QGridLayout()

        imageLabel = QLabel(self)
        pixmap = QPixmap("CPS_Logo_Col.png")
        pixmap = pixmap.scaled(
            300, 300, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        imageLabel.setPixmap(pixmap)
        layout.addWidget(imageLabel, 0, 0, 2, 1)

        textLabel1 = QLabel("This tool helps with importing ITU database for IAU CPS RAS database.\n\n"
                            "Program version: v0.1a\n\n"
                            "Initial version of this program parses the ITU database entries.\n\n", self)
        textLabel1.setWordWrap(True)
        layout.addWidget(textLabel1, 0, 1)

        textLabel2 = QLabel("For more information, contact via either:<br><br>"
                            "<a href='mailto:ras.database@cps.iau.org'>ras.database@cps.iau.org</a><br><br>"
                            "<a href='mailto:boris.sorokin@skao.int'>boris.sorokin@skao.int</a>", self)
        textLabel2.setWordWrap(True)
        textLabel2.setOpenExternalLinks(True)
        textLabel2.setTextFormat(Qt.RichText)
        layout.addWidget(textLabel2, 1, 1)

        self.setLayout(layout)


class InteractiveDatabase(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Interactive database window')
        self.parent = parent

        self.desired_width = 1600
        self.desired_height = 900
        self.initUI()
        # self.setWindowModality(Qt.ApplicationModal)

    def initUI(self):
        # Set the window properties. Assuming that we have parent above
        self.setWindowTitle(
            f'Interactive Database Tool connected to database {self.parent.database_version} published on {self.parent.database_date.date()}')
        self.setWindowIcon(QIcon('cps-logo-mono.ico'))

        # Center the window on the screen or parent

        self.parent.centerWindow(self, self.parent)
        self.parent.animateOpening(
            self, self.desired_width, self.desired_height)

        centralWidget = QWidget(self)
        self.setCentralWidget(centralWidget)
        layout = QGridLayout(centralWidget)

        self.tableWidget = QTableWidget()
        headers = ["ITU Notice ID", "ITU Administration code", "ITU Country code        ", "Station Name",
                   "Provision", "Date received", "Longitude", "Latitude"]
        self.load_country_codes()
        self.tableWidget.setColumnCount(len(headers))
        self.tableWidget.setHorizontalHeaderLabels(headers)
        self.tableWidget.setSortingEnabled(True)
        self.load_data()
        self.tableWidget.cellDoubleClicked.connect(
            self.openDatabaseEntryDetails)
        layout.addWidget(self.tableWidget, 0, 0, 8, 1)

        self.interactionPanel = QGroupBox('Interaction Panel')
        layout_interaction = QGridLayout(self.interactionPanel)
        self.displayNamesCheckbox = QCheckBox(
            "Display ITU Names instead of codes", self.interactionPanel)
        self.displayNamesCheckbox.stateChanged.connect(self.updateTableDisplay)
        checkboxContainer = QWidget()
        checkboxLayout = QHBoxLayout()
        checkboxLayout.setAlignment(Qt.AlignCenter)
        checkboxLayout.addWidget(self.displayNamesCheckbox)
        checkboxContainer.setLayout(checkboxLayout)

        layout_interaction.addWidget(checkboxContainer, 0, 0, 1, 1)
        self.showMapButton = QPushButton(
            "Show stations on the map", self.interactionPanel)
        self.showMapButton.clicked.connect(self.showStationsOnMap)
        layout_interaction.addWidget(self.showMapButton, 1, 0, 1, 1)

        saveCsvButton = QPushButton("Save as CSV", self.interactionPanel)
        layout_interaction.addWidget(saveCsvButton, 0, 1, 1, 1)
        saveCsvButton.clicked.connect(self.saveAsCsv)

        saveWordButton = QPushButton("Save as DOCX", self.interactionPanel)
        layout_interaction.addWidget(saveWordButton, 1, 1, 1, 1)
        saveWordButton.clicked.connect(self.saveAsWord)

        layout.addWidget(self.interactionPanel, 8, 0, 2, 1)

        self.statusBar().showMessage(
            f'Connected to database {self.parent.database_version} published on {self.parent.database_date.date()}')
        self.statusBar().setStyleSheet("""
            QStatusBar {
                background-color: #404040;
                color: white;
                border: 2px solid black;
            }
        """)

        self.updateTableDisplay()
        self.statusBar().showMessage('    Ready')
        self.show()
        self.setFocus()
        self.raise_()
        self.activateWindow()

    def load_data(self):
        SQL = "SELECT ntc_id, adm, ctry, stn_name, prov, d_rcv, long_dec, lat_dec FROM com_el WHERE ntc_type='R' ORDER BY adm asc, stn_name asc;"
        self.rows = self.parent.parse_database(SQL)

        self.tableWidget.setRowCount(len(self.rows))
        for row_num, row_data in enumerate(self.rows):
            for column_num, data in enumerate(row_data):
                if column_num == 5:
                    data = data.strftime("%Y-%m-%d")
                item = QTableWidgetItem(str(data))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.tableWidget.setItem(row_num, column_num, item)

        self.tableWidget.resizeColumnsToContents()

    def load_country_codes(self):
        self.country_codes = {}
        with open('geographical-areas.csv', mode='r', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                self.country_codes[row[0]] = row[1]

    def updateTableDisplay(self):
        displayNames = self.displayNamesCheckbox.isChecked()
        if displayNames:
            self.tableWidget.horizontalHeaderItem(
                1).setText("ITU Administration name")
            self.tableWidget.horizontalHeaderItem(
                2).setText("ITU Country name")
            self.statusBar().showMessage(
                '    View switched to administration and country name view')
        else:
            self.tableWidget.horizontalHeaderItem(
                1).setText("ITU Administration code")
            self.tableWidget.horizontalHeaderItem(
                2).setText("ITU Country code")
            self.statusBar().showMessage(
                '    View switched to administration and country code view')
        for row_num, row_data in enumerate(self.rows):
            admin_code = row_data[1]
            admin_name = self.country_codes.get(admin_code, "Unknown")
            display_text = admin_name if displayNames else admin_code
            item = QTableWidgetItem(display_text)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            tooltip_text = admin_name if not displayNames else None
            item.setToolTip(tooltip_text)
            self.tableWidget.setItem(row_num, 1, item)

            country_code = row_data[2]
            country_name = self.country_codes.get(country_code, "Unknown")
            display_text = country_name if displayNames else country_code
            item = QTableWidgetItem(display_text)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            tooltip_text = country_name if not displayNames else None
            item.setToolTip(tooltip_text)

            self.tableWidget.setItem(row_num, 2, item)

    def showStationsOnMap(self):
        station_data = []

        for row in range(self.tableWidget.rowCount()):
            raw_adm_info = self.tableWidget.item(row, 1).text()
            raw_country_info = self.tableWidget.item(row, 2).text()
            station_name = self.tableWidget.item(row, 3).text()
            longitude = self.tableWidget.item(row, 6).text()
            latitude = self.tableWidget.item(row, 7).text()

            if self.displayNamesCheckbox.isChecked():
                code = [key for key, value in self.country_codes.items()
                        if value == raw_adm_info]
                administration_info = f"{raw_adm_info} ({code[0]})" if code else raw_adm_info
                code = [key for key, value in self.country_codes.items()
                        if value == raw_country_info]
                country_info = f"{raw_country_info} ({code[0]})" if code else raw_country_info
            else:
                administration_name = self.country_codes.get(
                    raw_adm_info, "Unknown Country")
                administration_info = f"{administration_name} ({raw_adm_info})"
                country_name = self.country_codes.get(
                    raw_country_info, "Unknown Country")
                country_info = f"{country_name} ({raw_country_info})"

            try:
                station_data.append((station_name, administration_info, country_info, float(
                    latitude), float(longitude)))
            except:
                station_data.append((station_name, administration_info, country_info, (
                    latitude), (longitude)))

        self.parent.animateClosing(self)
        self.showMinimized()
        self.setEnabled(False)
        self.mapWindow = MapWindow(station_data, self)

    def openDatabaseEntryDetails(self, row, column):
        ntc_id_item = self.tableWidget.item(row, 0)
        station_name_item = self.tableWidget.item(row, 3)

        ntc_id = ntc_id_item.text() if ntc_id_item else "Unknown"
        station_name = station_name_item.text() if station_name_item else "Unknown"

        self.parent.animateClosing(self)
        self.showMinimized()
        self.setEnabled(False)
        self.detailsWindow = DatabaseEntryDetails(ntc_id, station_name, self)

    def saveAsCsv(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save as CSV", f"RAS_DB_OVERVIEW_CSV_{self.parent.database_version}_{self.parent.database_date.date()}", "CSV Files (*.csv)")
        if filePath:
            if os.path.basename(filePath) == 'geographical-areas.csv':
                QMessageBox.critical(
                    self, "Error", "This file is an important app file and cannot be overwritten.")
                return
            try:
                with open(filePath, 'w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.writer(file)
                    headers = [self.tableWidget.horizontalHeaderItem(
                        i).text() for i in range(self.tableWidget.columnCount())]
                    writer.writerow(headers)
                    for row in range(self.tableWidget.rowCount()):
                        rowData = [self.tableWidget.item(row, i).text() if self.tableWidget.item(
                            row, i) else '' for i in range(self.tableWidget.columnCount())]
                        writer.writerow(rowData)
            except Exception as e:
                QMessageBox.critical(self, "CSV saving Error",
                                     f"An error occurred while preparing the csv file:\n{e}")

    def saveAsWord(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save as DOCX", f"RAS_DB_OVERVIEW_DOCX_{self.parent.database_version}_{self.parent.database_date.date()}", "Word Files (*.docx)")
        if filePath:
            doc = docx.Document()
            section = doc.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = section.page_height, section.page_width

            doc.add_heading(
                f'Overview RAS information as per database {self.parent.database_version} published on {self.parent.database_date.date()}', level=1)

            table = doc.add_table(rows=1, cols=self.tableWidget.columnCount())
            header_cells = table.rows[0].cells

            for column in range(self.tableWidget.columnCount()):
                header = self.tableWidget.horizontalHeaderItem(column)
                if header:
                    header_cells[column].text = header.text()

            for row in range(self.tableWidget.rowCount()):
                row_cells = table.add_row().cells
                for col in range(self.tableWidget.columnCount()):
                    item = self.tableWidget.item(row, col)
                    row_cells[col].text = item.text() if item else ""

            try:
                doc.save(filePath)
            except Exception as e:
                QMessageBox.critical(self, "Docx saving Error",
                                     f"An error occurred while preparing the docx file:\n{e}")

    def closeEvent(self, event):
        if self.parent:
            self.parent.animateClosing(self)
            self.parent.showNormal()
            self.parent.animateOpening(
                self.parent, self.parent.desired_width, self.parent.desired_height)
            self.parent.setFocus()
            self.parent.raise_()
            self.parent.activateWindow()
            self.parent.setEnabled(True)


class MapWindow(QMainWindow):
    def __init__(self, station_data=None, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.station_data = station_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle(
            f'Radio astronomy stations on a map as per database {self.parent.parent.database_version} published on {self.parent.parent.database_date.date()}')
        self.setWindowIcon(QIcon('cps-logo-mono.ico'))

        self.parent.parent.centerWindow(self, self.parent)
        self.parent.parent.animateOpening(
            self, desired_width=1600, desired_height=900)

        centralWidget = QWidget()
        self.setCentralWidget(centralWidget)
        layout = QGridLayout()
        centralWidget.setLayout(layout)

        self.mapPanel = QGroupBox("Map Panel")
        mapLayout = QGridLayout()
        self.mapPanel.setLayout(mapLayout)

        self.browser = QWebEngineView()

        mapLayout.addWidget(self.browser, 0, 0)

        self.interactionPanel = QGroupBox("Interaction Panel")
        interactionLayout = QGridLayout()
        self.interactionPanel.setLayout(interactionLayout)

        saveScreenshotButton = QPushButton(
            "Save as Screenshot", self.interactionPanel)
        interactionLayout.addWidget(saveScreenshotButton, 0, 0)
        saveScreenshotButton.clicked.connect(self.saveScreenshot)

        saveHTMLButton = QPushButton("Save as HTML", self.interactionPanel)
        interactionLayout.addWidget(saveHTMLButton, 1, 0)
        saveHTMLButton.clicked.connect(self.saveHTML)

        layout.addWidget(self.mapPanel, 0, 0)
        layout.addWidget(self.interactionPanel, 1, 0)
        layout.setRowStretch(0, 8)
        layout.setRowStretch(1, 2)

        self.show()
        self.setFocus()
        self.raise_()
        self.activateWindow()

        if self.station_data is not None:
            self.browser.setHtml(self.generateMapHTML(self.station_data))

    def generateMapHTML(self, station_data):
        try:
            html_parts = []
            html_parts.append("""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Full Widget Leaflet Map for Radio Astronomy Stations Database Tool</title>
                <meta charset="utf-8" />
                <link 
                    rel="stylesheet" 
                    href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
                />
                <script 
                    src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js">
                </script>
                <style>
                    body {
                        padding: 0;
                        margin: 0;
                    }
                    html, body, #map {
                        height: 100%;
                        width: 100%;
                    }
                </style>
            </head>
            <body>
                <div id="map"></div>
            
                <script>
                    var map = L.map('map', {attributionControl: false}).setView([0, 0], 2);
                    var myAttrControl = L.control.attribution().addTo(map);
                    myAttrControl.setPrefix('<a href="https://leafletjs.com/">Leaflet</a>');
                    
                    mapLink = 
                        '<a href="http://openstreetmap.org">OpenStreetMap</a>';
                    L.tileLayer(
                        'http://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                        attribution: 'Map data by &copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>, under <a href="https://opendatacommons.org/licenses/odbl/">ODbL.</a>',
                        maxZoom: 18,
                        }).addTo(map);
            """)

            for name, adm, ctr, lat, lon in station_data:
                adm_escaped = adm.replace("'", "&#39;").replace('"', '&quot;')
                ctr_escaped = ctr.replace("'", "&#39;").replace('"', '&quot;')
                html_parts.append(
                    f"        L.marker([{lat}, {lon}]).addTo(map).bindPopup('<b>{name}</b><br>Region country code:<b>{ctr_escaped}</b><br>Responsible administration: <b>{adm_escaped}</b>');")

            html_parts.append("""
                </script>
            </body>
            </html>
            """)

            html = ''.join(html_parts)

        except Exception as e:
            html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Error</title>
        </head>
        <body>
            <p>Error generating map: {str(e)}</p>
        </body>
        </html>
        """
        return html

    def saveScreenshot(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save Screenshot", f"RAS_DB_MAP_{self.parent.parent.database_version}_{self.parent.parent.database_date.date()}", "PNG Files (*.png)")
        if filePath:
            self.browser.grab().save(filePath)

    def saveHTML(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save HTML", f"RAS_DB_MAP_{self.parent.parent.database_version}_{self.parent.parent.database_date.date()}", "HTML Files (*.html)")
        if filePath:
            def save_html(html):
                with open(filePath, "w", encoding='utf-8') as file:
                    file.write(html)
            self.browser.page().toHtml(save_html)

    def closeEvent(self, event):
        if self.parent:
            self.parent.parent.animateClosing(self)
            self.parent.showNormal()
            self.parent.parent.animateOpening(
                self.parent, self.parent.desired_width, self.parent.desired_height)
            self.parent.setFocus()
            self.parent.raise_()
            self.parent.activateWindow()
            self.parent.setEnabled(True)


class DatabaseEntryDetails(QMainWindow):
    def __init__(self, ntc_id=None, station_name=None, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.ntc_id = ntc_id
        self.station_name = station_name
        self.station_rows = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle(
            f'Details of station {self.station_name} with notice id {self.ntc_id}')
        self.setWindowIcon(QIcon('cps-logo-mono.ico'))

        # Center the window on the screen or parent

        self.parent.parent.centerWindow(self, self.parent)
        self.parent.parent.animateOpening(
            self, desired_width=1600, desired_height=900)

        centralWidget = QWidget(self)
        self.setCentralWidget(centralWidget)
        layout = QGridLayout(centralWidget)

        stationInfoPanel = QGroupBox('Station Information')
        self.stationInfoLayout = QGridLayout(stationInfoPanel)
        self.stationInfoTable = QTableWidget()
        headers = ['Notice ID', 'Longitude (Degrees)', 'East/West', 'Minutes', 'Seconds', 'Latitude (Degrees)',
                   'North/South', 'Minutes', 'Seconds', 'Min elevation', 'Max elevation', 'Min azimuth', 'Max Azimuth',
                   'Antenna altitude, m']
        self.stationInfoTable.setColumnCount(len(headers))
        self.stationInfoTable.setHorizontalHeaderLabels(headers)
        self.stationInfoTable.setSortingEnabled(True)

        self.stationInfoLayout.addWidget(self.stationInfoTable)
        layout.addWidget(stationInfoPanel, 0, 0, 1, 4)

        self.beamInfoPanel = QGroupBox('Beams Information')
        self.beamInfoLayout = QGridLayout(self.beamInfoPanel)
        self.beamInfoTable = QTableWidget()
        headers = ['Beam name', 'Antenna Code', 'Antenna diameter, m', 'Antenna gain, dBi',
                   'Noise temp, K', 'Frequency minimum, MHz', 'Frequency maximum, MHz', 'Centre frequency, MHz']
        self.beamInfoTable.setColumnCount(len(headers))
        self.beamInfoTable.setHorizontalHeaderLabels(headers)
        self.beamInfoTable.setSortingEnabled(True)

        self.beamInfoLayout.addWidget(self.beamInfoTable)
        layout.addWidget(self.beamInfoPanel, 1, 0, 1, 4)

        self.controlPanel = QGroupBox('Interaction Panel')
        self.controlLayout = QGridLayout(self.controlPanel)
        layout.addWidget(self.controlPanel, 2, 0, 1, 1)

        openAntennaTableButton = QPushButton(
            "Open Antenna Code Table (System PDF viewer)", self.controlPanel)
        self.controlLayout.addWidget(openAntennaTableButton, 0, 0)
        openAntennaTableButton.clicked.connect(self.openAntennaTable)

        saveCsvButton = QPushButton("Save Tables to CSV", self.controlPanel)
        self.controlLayout.addWidget(saveCsvButton)
        saveCsvButton.clicked.connect(self.saveTablesToCsv)

        saveDocxButton = QPushButton("Save Tables to DOCX", self.controlPanel)
        self.controlLayout.addWidget(saveDocxButton)
        saveDocxButton.clicked.connect(self.saveTablesToWord)

        self.mapPanel = QGroupBox('Map Panel')
        self.mapLayout = QGridLayout(self.mapPanel)

        self.browser = QWebEngineView()

        layout.addWidget(self.mapPanel, 2, 1, 1, 3)

        self.mapLayout.addWidget(self.browser, 0, 0)

        self.load_data()
        self.browser.setHtml(self.generateMapHTML())

        layout.setRowStretch(0, 1)
        layout.setRowStretch(1, 4)
        layout.setRowStretch(2, 5)

        self.show()
        self.setFocus()
        self.raise_()
        self.activateWindow()

    def load_data(self):
        SQL = f"SELECT long_deg, long_ew, long_min, long_sec, lat_deg, lat_ns, lat_min, lat_sec, elev_min, elev_max, azm_fr, azm_to, ant_alt FROM e_stn WHERE ntc_id={str(self.ntc_id)};"
        self.station_rows = self.parent.parent.parse_database(SQL)

        self.stationInfoTable.setRowCount(len(self.station_rows))
        for row_num, row_data in enumerate(self.station_rows):
            item = QTableWidgetItem(str(self.ntc_id))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.stationInfoTable.setItem(row_num, 0, item)

            for column_num, data in enumerate(row_data):
                if data == None:
                    data = 'N/A'
                item = QTableWidgetItem(str(data))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.stationInfoTable.setItem(row_num, column_num+1, item)

        self.stationInfoTable.resizeColumnsToContents()

        SQL = f"SELECT beam_name, pattern_id, ant_diam, gain FROM e_ant WHERE ntc_id={str(self.ntc_id)};"
        self.beam_rows = self.parent.parent.parse_database(SQL)

        self.beamInfoTable.setRowCount(0)
        for beam_index, beam_row_data in enumerate(self.beam_rows):
            beam_name = beam_row_data[0]
            SQL = f"SELECT noise_t, freq_min, freq_max FROM grp WHERE ntc_id={str(self.ntc_id)} AND beam_name='{beam_name}';"
            self.grp_rows = self.parent.parent.parse_database(SQL)

            SQL = f"SELECT freq_mhz FROM freq WHERE ntc_id={str(self.ntc_id)} AND beam_name='{beam_name}';"
            self.freq_rows = self.parent.parent.parse_database(SQL)

            if not beam_row_data[1] == None:
                SQL = f"SELECT pattern FROM ant_type WHERE pattern_id={beam_row_data[1]};"
                try:
                    antenna_code = self.parent.parent.parse_database(SQL)[0][0]
                except:
                    try:
                        antenna_code = self.parent.parent.parse_database(SQL)[
                            0]
                    except:
                        antenna_code = None
                if antenna_code == None:
                    antenna_code = 'N/A'
            else:
                antenna_code = 'N/A'

            for grp_index, grp_row_data in enumerate(self.grp_rows):
                current_row_count = self.beamInfoTable.rowCount()
                self.beamInfoTable.insertRow(current_row_count)
                freq_row_data = self.freq_rows[grp_index][0]

                combined_row_data = list(
                    beam_row_data) + list(grp_row_data) + [freq_row_data]

                for column_num, data in enumerate(combined_row_data):
                    if data == None:
                        data = 'N/A'
                    if column_num == 1:
                        data = antenna_code
                    item = QTableWidgetItem(str(data))
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.beamInfoTable.setItem(
                        current_row_count, column_num, item)

        self.beamInfoTable.resizeColumnsToContents()

    def generateMapHTML(self):
        SQL = f"SELECT adm, ctry, stn_name, long_dec, lat_dec FROM com_el WHERE ntc_id={str(self.ntc_id)};"
        self.rows = self.parent.parent.parse_database(SQL)[0]

        adm = self.rows[0]
        ctr = self.rows[1]

        administration_name = self.parent.country_codes.get(
            adm, "Unknown Country")
        adm = f"{administration_name} ({adm})"
        country_name = self.parent.country_codes.get(ctr, "Unknown Country")
        ctr = f"{country_name} ({ctr})"

        name = self.rows[2]
        lon = float(self.rows[3])
        lat = float(self.rows[4])

        try:
            html_parts = []
            html_parts.append("""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Full Widget Leaflet Map for Radio Astronomy Stations Database Tool</title>
                <meta charset="utf-8" />
                <link 
                    rel="stylesheet" 
                    href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
                />
                <script 
                    src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js">
                </script>
                <style>
                    body {
                        padding: 0;
                        margin: 0;
                    }
                    html, body, #map {
                        height: 100%;
                        width: 100%;
                    }
                </style>
            </head>
            <body>
                <div id="map"></div>
            
                <script>
                """)
            html_parts.append(f"""
                    var map = L.map('map', {{attributionControl: false}}).setView([{lat}, {lon}], 5);
                    """)
            html_parts.append("""
                    var myAttrControl = L.control.attribution().addTo(map);
                    myAttrControl.setPrefix('<a href="https://leafletjs.com/">Leaflet</a>');
                    
                    mapLink = 
                        '<a href="http://openstreetmap.org">OpenStreetMap</a>';
                    L.tileLayer(
                        'http://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                        attribution: 'Map data by &copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>, under <a href="https://opendatacommons.org/licenses/odbl/">ODbL.</a>',
                        maxZoom: 18,
                        }).addTo(map);
            """)
            adm_escaped = adm.replace("'", "&#39;").replace('"', '&quot;')
            ctr_escaped = ctr.replace("'", "&#39;").replace('"', '&quot;')
            html_parts.append(
                f"        L.marker([{lat}, {lon}]).addTo(map).bindPopup('<b>{name}</b><br>Region country code:<b>{ctr_escaped}</b><br>Responsible administration: <b>{adm_escaped}</b>').openPopup();")
            html_parts.append("""
                </script>
            </body>
            </html>
            """)

            html = ''.join(html_parts)

        except Exception as e:
            html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Error</title>
        </head>
        <body>
            <p>Error generating map: {str(e)}</p>
        </body>
        </html>
        """
        return html

    def openAntennaTable(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile('Table6toPreface.pdf'))

    def saveTablesToCsv(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save as CSV", f"RAS_DB_CSV_{self.station_name}_{self.parent.parent.database_version}_{self.parent.parent.database_date.date()}", "CSV Files (*.csv)")
        if filePath:
            if os.path.basename(filePath) == 'geographical-areas.csv':
                QMessageBox.critical(
                    self, "Error", "This file is an important app file and cannot be overwritten.")
                return
            try:
                with open(filePath, 'w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.writer(file)

                    headers = [self.stationInfoTable.horizontalHeaderItem(
                        i).text() for i in range(self.stationInfoTable.columnCount())]
                    writer.writerow(headers)
                    for row in range(self.stationInfoTable.rowCount()):
                        rowData = [self.stationInfoTable.item(row, i).text() if self.stationInfoTable.item(
                            row, i) else '' for i in range(self.stationInfoTable.columnCount())]
                        writer.writerow(rowData)

                    writer.writerow([])

                    headers = [self.beamInfoTable.horizontalHeaderItem(
                        i).text() for i in range(self.beamInfoTable.columnCount())]
                    writer.writerow(headers)
                    for row in range(self.beamInfoTable.rowCount()):
                        rowData = [self.beamInfoTable.item(row, i).text() if self.beamInfoTable.item(
                            row, i) else '' for i in range(self.beamInfoTable.columnCount())]
                        writer.writerow(rowData)
            except Exception as e:
                QMessageBox.critical(self, "CSV saving Error",
                                     f"An error occurred while preparing the csv file:\n{e}")

    def saveTablesToWord(self):
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save as DOCX", f"RAS_DB_DOCX_{self.station_name}_{self.parent.parent.database_version}_{self.parent.parent.database_date.date()}", "Word Files (*.docx)")
        if filePath:
            doc = docx.Document()
            section = doc.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = section.page_height, section.page_width

            doc.add_heading(
                f'Station "{self.station_name}" information as per database {self.parent.parent.database_version} published on {self.parent.parent.database_date.date()}', level=1)
            doc.add_heading("General information", level=2)
            SQL = f"SELECT adm, ctry, long_dec, lat_dec FROM com_el WHERE ntc_id={self.ntc_id};"
            rows = self.parent.parent.parse_database(SQL)[0]
            adm = rows[0]
            administration_name = self.parent.country_codes.get(
                adm, "Unknown Country")
            doc.add_paragraph(
                f"Responsible administration: {administration_name}")
            ctry = rows[1]
            country_name = self.parent.country_codes.get(
                ctry, "Unknown Country")
            doc.add_paragraph(f"Region country code: {country_name}")
            long_dec = rows[2]
            doc.add_paragraph(
                f"Longitude {abs(long_dec)} {self.station_rows[0][1]}")
            lat_dec = rows[3]
            doc.add_paragraph(
                f"Latitude {abs(lat_dec)} {self.station_rows[0][5]}")

            doc.add_heading("Station Information Table", level=2)

            doc.add_paragraph("")
            table = doc.add_table(
                rows=1, cols=self.stationInfoTable.columnCount())
            header_cells = table.rows[0].cells

            for col in range(self.stationInfoTable.columnCount()):
                header_cells[col].text = self.stationInfoTable.horizontalHeaderItem(
                    col).text()
            for row in range(self.stationInfoTable.rowCount()):
                row_cells = table.add_row().cells
                for col in range(self.stationInfoTable.columnCount()):
                    row_cells[col].text = self.stationInfoTable.item(
                        row, col).text() if self.stationInfoTable.item(row, col) else ''

            doc.add_paragraph("")

            doc.add_heading("Beams Information Table", level=2)

            doc.add_paragraph("")

            table = doc.add_table(
                rows=1, cols=self.beamInfoTable.columnCount())
            header_cells = table.rows[0].cells

            for col in range(self.beamInfoTable.columnCount()):
                header_cells[col].text = self.beamInfoTable.horizontalHeaderItem(
                    col).text()
            for row in range(self.beamInfoTable.rowCount()):
                row_cells = table.add_row().cells
                for col in range(self.beamInfoTable.columnCount()):
                    row_cells[col].text = self.beamInfoTable.item(
                        row, col).text() if self.beamInfoTable.item(row, col) else ''

            doc.add_paragraph("")

            try:
                doc.save(filePath)
            except Exception as e:
                QMessageBox.critical(self, "Docx saving Error",
                                     f"An error occurred while preparing the docx file:\n{e}")

    def closeEvent(self, event):
        if self.parent:
            self.parent.parent.animateClosing(self)
            self.parent.showNormal()
            self.parent.parent.animateOpening(
                self.parent, self.parent.desired_width, self.parent.desired_height)
            self.parent.setFocus()
            self.parent.raise_()
            self.parent.activateWindow()
            self.parent.setEnabled(True)


if __name__ == '__main__':
    qt_app = QApplication(sys.argv)
    qt_gui = MainApp()
    sys.exit(qt_app.exec_())
