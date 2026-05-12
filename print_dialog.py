import sys
if 'linux' not in sys.platform:
    import win32print
    import wmi
else:
    import cups
from PyQt6.QtCore import QSize, QRectF, Qt, QMarginsF, QRect
from PyQt6.QtGui import QPixmap, QIcon, QPainter, QPageLayout, QPageSize, QPen, QBrush, QImage
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtWidgets import QWidget, QLabel, QPushButton, QHBoxLayout, QComboBox, QDialog, QVBoxLayout, \
    QMessageBox, QRadioButton, QButtonGroup, QToolButton

from widgets import AutoSelectSpinBox, AutoSelectLineEdit


class PrintDialog(QDialog):
    """
    Class implementing QDialog to show the user a print dialog, also showing a preview of the item to be printed.
    :param QPdfDocument pdf_doc: The QPdfDocument to be printed
    :param GUI gui: the current instance of a gui
    :param boolean landscape: optional: whether the pdf is to be printed in landscape orientation
    """
    def __init__(self, pdf_doc, gui, landscape=False, page_margins_horizontal=0.5, page_margins_vertical=0.5):
        """
        Class implementing QDialog to show the user a print dialog, also showing a preview of the item to be printed.
        :param QPdfDocument pdf_doc: a QPdfDocument
        :param GUI gui: the current instance of a gui
        :param boolean landscape: optional: whether the pdf is to be printed in landscape orientation
        :param float page_margins_horizontal: optional: the left and right margins, in inches
        :param float page_margins_vertical: optional: the top and bottom margins, in inches
        """
        super().__init__()
        self.gui = gui
        self.landscape = landscape
        self.page_margins_horizontal = page_margins_horizontal
        self.page_margins_vertical = page_margins_vertical
        self.printer_properties = {'name': '', 'printable_rect_inch': QRectF(), 'printable_rect_px': QRectF(), 'printable_margins': QMarginsF() }

        self.pdf = pdf_doc
        self.num_pages = self.pdf.pageCount()
        self.pages = []
        self.current_page = 0

        self.layout = QHBoxLayout(self)

        self.setWindowTitle('Print')
        self.setGeometry(50, 50, 100, 100)
        self.init_components()
        self.get_printer_properties()
        self.get_pages()
        self.draw_preview(self.pages[0])

    def init_components(self):
        """
        Creates the widgets to be shown in the dialog
        """
        preview_widget = QWidget()
        preview_widget.setObjectName('preview_widget')
        preview_widget.setStyleSheet('#preview_widget { background: white; border-right: 2px solid black; }')
        preview_layout = QVBoxLayout(preview_widget)
        self.layout.addWidget(preview_widget)

        self.preview_label = QLabel('Preview:')
        self.preview_label.setFont(self.gui.bold_font)
        self.preview_label.setAutoFillBackground(False)
        preview_layout.addWidget(self.preview_label)

        self.pdf_label = QLabel()
        preview_layout.addWidget(self.pdf_label)

        nav_button_widget = QWidget()
        nav_button_layout = QHBoxLayout(nav_button_widget)
        preview_layout.addWidget(nav_button_widget)
        preview_layout.addStretch()

        previous_button = QPushButton()
        previous_button.setIcon(QIcon('resources/previous.svg'))
        previous_button.setAutoFillBackground(False)
        previous_button.setStyleSheet('border: none')
        previous_button.clicked.connect(self.previous_page)
        nav_button_layout.addStretch()
        nav_button_layout.addWidget(previous_button)
        nav_button_layout.addSpacing(20)

        self.page_label = QLabel('Page 1 of ' + str(self.num_pages))
        self.page_label.setFont(self.gui.bold_font)
        nav_button_layout.addWidget(self.page_label)
        nav_button_layout.addSpacing(20)

        next_button = QPushButton()
        next_button.setIcon(QIcon('resources/next.svg'))
        next_button.setAutoFillBackground(False)
        next_button.setStyleSheet('border: none')
        next_button.clicked.connect(self.next_page)
        nav_button_layout.addWidget(next_button)
        nav_button_layout.addStretch()

        print_options_widget = QWidget()
        print_options_layout = QVBoxLayout(print_options_widget)
        print_options_layout.setSpacing(0)
        self.layout.addWidget(print_options_widget)

        print_button = QToolButton()
        print_button.setText('Print')
        print_button.setIcon(QIcon('resources/print.svg'))
        print_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        print_button.setFont(self.gui.bold_font)
        print_button.setIconSize(QSize(64, 64))
        print_button.setFixedWidth(72)
        print_button.clicked.connect(self.do_print)
        print_options_layout.addWidget(print_button)
        print_options_layout.addSpacing(20)

        printer_widget = QWidget()
        printer_layout = QHBoxLayout(printer_widget)
        printer_layout.setContentsMargins(0, 0, 0, 0)
        print_options_layout.addWidget(printer_widget)
        print_options_layout.addSpacing(10)

        printer_label = QLabel('Printer')
        printer_label.setFont(self.gui.bold_font)
        printer_layout.addWidget(printer_label)

        self.printer_combobox = self.get_printers()
        self.printer_combobox.setMinimumHeight(40)
        self.printer_combobox.setFont(self.gui.standard_font)
        self.printer_combobox.currentIndexChanged.connect(self.get_printer_properties)
        printer_layout.addWidget(self.printer_combobox)

        copies_widget = QWidget()
        copies_layout = QHBoxLayout(copies_widget)
        copies_layout.setContentsMargins(0, 0, 0, 0)
        print_options_layout.addWidget(copies_widget)
        print_options_layout.addSpacing(10)

        copies_label = QLabel('Copies')
        copies_label.setFont(self.gui.bold_font)
        copies_layout.addWidget(copies_label)
        copies_layout.addStretch()

        self.copies_spinbox = AutoSelectSpinBox()
        self.copies_spinbox.setMinimumSize(100, 40)
        self.copies_spinbox.setRange(1, 500)
        self.copies_spinbox.setValue(1)
        self.copies_spinbox.setFont(self.gui.standard_font)
        copies_layout.addWidget(self.copies_spinbox)

        orientation_widget = QWidget()
        orientation_widget.setMinimumHeight(40)
        orientation_layout = QHBoxLayout(orientation_widget)
        orientation_layout.setContentsMargins(0, 0, 0, 0)
        print_options_layout.addWidget(orientation_widget)
        print_options_layout.addSpacing(10)

        orientation_label = QLabel('Orientation')
        orientation_label.setFont(self.gui.bold_font)
        orientation_layout.addWidget(orientation_label)
        orientation_layout.addStretch()

        portrait_radio_button = QRadioButton('Portrait')
        portrait_radio_button.setFont(self.gui.standard_font)
        orientation_layout.addWidget(portrait_radio_button)

        landscape_radio_button = QRadioButton('Landscape')
        landscape_radio_button.setFont(self.gui.standard_font)
        orientation_layout.addWidget(landscape_radio_button)

        if self.landscape:
            landscape_radio_button.setChecked(True)
            portrait_radio_button.setChecked(False)
        else:
            landscape_radio_button.setChecked(False)
            portrait_radio_button.setChecked(True)

        self.orientation_group = QButtonGroup()
        self.orientation_group.setExclusive(True)
        self.orientation_group.buttonClicked.connect(self.orientation_changed)
        self.orientation_group.addButton(portrait_radio_button, 0)
        self.orientation_group.addButton(landscape_radio_button, 1)

        duplex_widget = QWidget()
        print_options_layout.addWidget(duplex_widget)
        print_options_layout.addSpacing(10)
        duplex_layout = QHBoxLayout(duplex_widget)
        duplex_layout.setContentsMargins(0, 0, 0, 0)

        duplex_label = QLabel('2-Sided Printing:')
        duplex_label.setFont(self.gui.bold_font)
        duplex_layout.addWidget(duplex_label)
        duplex_layout.addStretch()

        self.duplex_combobox = QComboBox()
        self.duplex_combobox.setFont(self.gui.standard_font)
        self.duplex_combobox.setMinimumHeight(40)
        self.duplex_combobox.addItem(QIcon('resources/one_sided.svg'), 'Print 1-Sided')
        self.duplex_combobox.addItem(QIcon('resources/flip_long_edge.svg'), 'Flip on Long Edge')
        self.duplex_combobox.addItem(QIcon('resources/flip_short_edge.svg'), 'Flip on Short Edge')
        self.duplex_combobox.setIconSize(QSize(36, 36))
        self.duplex_combobox.setStyleSheet('QListView::item { height: 40px; }')
        duplex_layout.addWidget(self.duplex_combobox)

        range_widget = QWidget()
        range_layout = QHBoxLayout(range_widget)
        range_layout.setContentsMargins(0, 0, 0, 0)
        print_options_layout.addWidget(range_widget)

        range_label = QLabel('Pages')
        range_label.setFont(self.gui.bold_font)
        range_layout.addWidget(range_label)
        range_layout.addStretch()

        self.range_line_edit = AutoSelectLineEdit()
        self.range_line_edit.setMinimumHeight(40)
        self.range_line_edit.setFont(self.gui.standard_font)
        range_layout.addWidget(self.range_line_edit)

        details_label = QLabel('(i.e. "2", "1-3", "1, 3", or leave blank to print all pages)')
        details_label.setFont(self.gui.standard_font)
        details_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        print_options_layout.addWidget(details_label, Qt.AlignmentFlag.AlignRight)
        print_options_layout.addStretch()

    def calculate_preview_rects(self):
        if self.landscape:
            ratio = 600 / min(self.pdf.pagePointSize(0).width(), self.pdf.pagePointSize(0).height())
            preview_rect = QRect(
                0,
                0,
                600,
                int(600 * 8.5 / 11)
            )
            preview_print_area_rect = QRect(
                int(self.page_margins_horizontal * ratio),
                int(self.page_margins_vertical * ratio),
                preview_rect.width() - int(self.page_margins_horizontal * ratio * 2),
                preview_rect.height() - int(self.page_margins_vertical * ratio * 2)
            )
        else:
            ratio = 600 / max(self.pdf.pagePointSize(0).width(), self.pdf.pagePointSize(0).height())
            preview_rect = QRect(
                0,
                0,
                600,
                int(600 * 11 / 8.5)
            )
            preview_print_area_rect = QRect(
                int(self.page_margins_horizontal * ratio),
                int(self.page_margins_vertical * ratio),
                preview_rect.width() - int(self.page_margins_horizontal * ratio * 2),
                preview_rect.height() - int(self.page_margins_vertical * ratio * 2)
            )

        return preview_rect, preview_print_area_rect

    def get_printers(self):
        """
        Obtain a list of printers currently installed on the system
        """
        if not 'linux' in sys.platform:
            win_management = wmi.WMI()
            printers = win_management.Win32_Printer()
        else:
            connection = cups.Connection()
            printers = connection.getPrinters()

        printer_combobox = QComboBox()
        default = ''

        if not 'linux' in sys.platform:
            for printer in printers:
                if not printer.Hidden:
                    printer_combobox.addItem(printer.Name)
                if printer.Default:
                    default = printer.Name
        else:
            count = 0
            for printer in printers.keys():
                printer_combobox.addItem(printer)
                if count == 0:
                    default = printer
                count += 1

        printer_combobox.setCurrentText(default)
        return printer_combobox

    def get_printer_properties(self):
        prn = self.printer_combobox.currentText()
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPrinterName(prn)
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.Letter))

        page_rect_inch = printer.pageRect(QPrinter.Unit.Inch)
        dpi_x = printer.physicalDpiX()
        dpi_y = printer.physicalDpiY()
        unprintable_margin_x = (8.5 - page_rect_inch.width()) / 2
        unprintable_margin_y = (11 - page_rect_inch.height()) / 2

        self.printer_properties['name'] = prn
        self.printer_properties['letter_page_size_px'] = QRectF(
            0,
            0,
            8.5 * dpi_x,
            11 * dpi_y
        )
        self.printer_properties['printable_rect_inch'] = QRectF(
            unprintable_margin_x,
            unprintable_margin_y,
            page_rect_inch.width(),
            page_rect_inch.height()
        )
        self.printer_properties['printable_rect_px'] = QRectF(
            unprintable_margin_x * dpi_x,
            unprintable_margin_y * dpi_y,
            page_rect_inch.width() * dpi_x,
            page_rect_inch.height() * dpi_y
        )
        self.printer_properties['printable_margins'] = QMarginsF(
            self.printer_properties['printable_rect_px'].x() - self.printer_properties['letter_page_size_px'].x(),
            self.printer_properties['printable_rect_px'].y() - self.printer_properties['letter_page_size_px'].y(),
            self.printer_properties['letter_page_size_px'].width() - (self.printer_properties['printable_rect_px'].x() + self.printer_properties['printable_rect_px'].width()),
            self.printer_properties['letter_page_size_px'].height() - (self.printer_properties['printable_rect_px'].y() + self.printer_properties['printable_rect_px'].height())
        )
        max_x_margin = max(
            self.printer_properties['printable_margins'].left(), self.printer_properties['printable_margins'].right())
        max_y_margin = max(
            self.printer_properties['printable_margins'].top(), self.printer_properties['printable_margins'].bottom())
        self.printer_properties['centered_printable_rect_px'] = QRectF(
            max_x_margin,
            max_y_margin,
            self.printer_properties['letter_page_size_px'].width() - (max_x_margin * 2),
            self.printer_properties['letter_page_size_px'].height() - (max_y_margin * 2)
        )

    def get_pages(self):
        """
        Gets the individual pages of the pdf file and converts each to a QPixmap
        """
        self.pages = []
        for i in range(self.pdf.pageCount()):
            preview_rect, preview_print_area_rect = self.calculate_preview_rects()
            image = self.pdf.render(
                i,
                QSize(
                    preview_print_area_rect.width(),
                    preview_print_area_rect.height()
                )
            )
            self.pages.append(image)

    def draw_preview(self, image):
        self.pdf_label.clear()

        preview_rect, preview_print_area_rect = self.calculate_preview_rects()
        pixmap = QPixmap(
            preview_rect.width() + 4,
            preview_rect.height() + 4
        )
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # draw the shadow
        pen = QPen(Qt.GlobalColor.lightGray)
        brush = QBrush(Qt.GlobalColor.lightGray)
        painter.setPen(pen)
        painter.setBrush(brush)
        painter.drawRect(
            4,
            4,
            preview_rect.width(),
            preview_rect.height()
        )

        # draw the white paper
        pen = QPen(Qt.GlobalColor.black)
        pen.setWidth(2)
        brush = QBrush(Qt.GlobalColor.white)
        painter.setPen(pen)
        painter.setBrush(brush)
        painter.drawRect(
            0,
            0,
            preview_rect.width(),
            preview_rect.height()
        )

        painter.drawImage(preview_print_area_rect.x(), preview_print_area_rect.y(), image)
        painter.end()
        
        self.pdf_label.setPixmap(pixmap)

    def previous_page(self):
        """
        Method to show the previous page in the PDF file upon user input
        """
        if not self.current_page == 0:
            self.current_page -= 1
            self.page_label.setText('Page ' + str(self.current_page + 1) + ' of ' + str(self.num_pages))
            self.draw_preview(self.pages[self.current_page])

    def next_page(self):
        """
        Method to show the next page in the PDF file upon user input
        """
        if not self.current_page == self.num_pages - 1:
            self.current_page += 1
            self.page_label.setText('Page ' + str(self.current_page + 1) + ' of ' + str(self.num_pages))
            self.draw_preview(self.pages[self.current_page])

    def orientation_changed(self):
        if self.orientation_group.checkedId() == 0:
            self.landscape = False
        else:
            self.landscape = True

        self.get_pages()
        self.draw_preview(self.pages[self.current_page])

    def parse_page_range(self, page_range):
        pages_to_print = []
        if ',' in page_range:
            page_range_split = page_range.split(',')
            for page in page_range_split:
                if '-' in page:
                    this_split = page.strip().split('-')
                    if not len(this_split) == 2:
                        return -1

                    try:
                        low = int(this_split[0].strip()) - 1
                        high = int(this_split[1].strip()) - 1
                    except ValueError:
                        return -1

                    if low > high:
                        low = int(this_split[1].strip()) - 1
                        high = int(this_split[0].strip()) - 1
                    for i in range(low, high + 1):
                        pages_to_print.append(i)
                else:
                    try:
                        pages_to_print.append(int(page.strip()) - 1)
                    except ValueError:
                        return -1
        elif '-' in page_range:
            this_split = page_range.strip().split('-')
            if not len(this_split) == 2:
                return -1

            try:
                low = int(this_split[0].strip()) - 1
                high = int(this_split[1].strip()) - 1

                if low > high:
                    low = int(this_split[1].strip()) - 1
                    high = int(this_split[0].strip()) - 1
                for i in range(low, high + 1):
                    pages_to_print.append(i)
            except ValueError:
                return -1
        else:
            try:
                pages_to_print.append(int(page_range.strip()) - 1)
            except ValueError:
                return -1

        return pages_to_print

    def do_print(self):
        """if not 'linux' in sys.platform:
            if self.orientation_group.button(1).isChecked():
                orientation = 2
            else:
                orientation = 1

            defaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}
            printer_handle = win32print.OpenPrinter(self.printer_properties['name'], defaults)
            properties = win32print.GetPrinter(printer_handle, 2)
            original_orientation = properties['pDevMode'].Orientation
            original_duplex = properties['pDevMode'].Duplex

            properties['pDevMode'].Orientation = orientation
            properties['pDevMode'].Duplex = self.duplex_combobox.currentIndex() + 1
            win32print.SetPrinter(printer_handle, 2, properties, 0)"""

        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPrinterName(self.printer_properties['name'])
        printer.setDocName('Bulletin Builder')
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.Letter))

        if self.orientation_group.button(1).isChecked():
            printer.setPageOrientation(QPageLayout.Orientation.Landscape)
        else:
            printer.setPageOrientation(QPageLayout.Orientation.Portrait)

        if self.duplex_combobox.currentIndex() == 0:
            printer.setDuplex(QPrinter.DuplexMode.DuplexNone)
        elif self.duplex_combobox.currentIndex() == 1:
            printer.setDuplex(QPrinter.DuplexMode.DuplexLongSide)
        elif self.duplex_combobox.currentIndex() == 2:
            printer.setDuplex(QPrinter.DuplexMode.DuplexShortSide)

        page_range = self.range_line_edit.text()
        pages_to_print = []
        if len(page_range.strip()) == 0:
            for i in range(self.num_pages):
                pages_to_print.append(i)
        else:
            pages_to_print = self.parse_page_range(page_range)

        if pages_to_print == -1:
            QMessageBox.information(
                self.gui,
                'Invalid Page Range',
                'Unable to determine pages to print from given page range(s).',
                QMessageBox.StandardButton.Ok
            )
            return

        dpi_x = printer.logicalDpiX()
        dpi_y = printer.logicalDpiY()

        # Ignore printable area and use the page rect to determine print rects
        page_rect_inch_f = printer.pageRect(QPrinter.Unit.Inch)
        render_rect_f = QRectF(
            0, # printer won't begin printing until after its own margins have been advanced, so x and y are inferred
            0,
            (page_rect_inch_f.width() + page_rect_inch_f.x()) * dpi_x, # rect is already offset by the internal margins so add only the second margin
            (page_rect_inch_f.height() + page_rect_inch_f.y()) * dpi_y
        )

        self.renders = []
        for page in pages_to_print:
            image = self.pdf.render(
                page,
                QSize(
                    int(render_rect_f.width()),
                    int(render_rect_f.height()),
                )
            )
            self.renders.append(image)

        printer.setCopyCount(self.copies_spinbox.value())
        painter = QPainter(printer)
        for i in range(len(self.renders)):
            if i > 0:
                printer.newPage()

            painter.drawImage(
                render_rect_f,
                self.renders[i]
            )
        painter.end()

        self.done(0)

    def cancel(self):
        """
        Method to end the QDialog upon user clicking the "Cancel" button
        """
        self.done(0)

    def closeEvent(self, evt):
        self.done(0)
