import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QComboBox, QMessageBox, QFileDialog
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from calendar import monthrange
from datetime import datetime, timedelta
import random

class ExcelGenerator(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Excel Tablo Oluşturucu')
        self.setGeometry(100, 100, 600, 400)
        
        self.label_yil = QLabel('Yıl Seçin:')
        self.combo_yil = QComboBox()
        self.combo_yil.addItems([str(yil) for yil in range(2020, 2031)])
        
        self.label_ay = QLabel('Ay Seçin:')
        self.combo_ay = QComboBox()
        self.combo_ay.addItems([
            'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran',
            'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'
        ])
        
        self.label_gun_basi_km = QLabel('Gün Başı Kilometre:')
        self.input_gun_basi_km = QLineEdit()
        
        self.label_km_aralik = QLabel('Yapılan Kilometre Aralığı (örn: 100-300):')
        self.input_km_aralik = QLineEdit()
        
        self.label_gorev_yeri = QLabel('Göreve Gidilen Yer:')
        self.input_gorev_yeri = QLineEdit()
        
        self.label_haftasonu = QLabel('Hafta Sonu Çalışma Durumu:')
        self.combo_haftasonu = QComboBox()
        self.combo_haftasonu.addItems(['Çalışıyor', 'Çalışmıyor'])
        
        self.button_kaydet = QPushButton('Kaydet ve Listele')
        self.button_kaydet.clicked.connect(self.kaydet_ve_listele)
        
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.label_yil)
        self.layout.addWidget(self.combo_yil)
        self.layout.addWidget(self.label_ay)
        self.layout.addWidget(self.combo_ay)
        self.layout.addWidget(self.label_gun_basi_km)
        self.layout.addWidget(self.input_gun_basi_km)
        self.layout.addWidget(self.label_km_aralik)
        self.layout.addWidget(self.input_km_aralik)
        self.layout.addWidget(self.label_gorev_yeri)
        self.layout.addWidget(self.input_gorev_yeri)
        self.layout.addWidget(self.label_haftasonu)
        self.layout.addWidget(self.combo_haftasonu)
        self.layout.addWidget(self.button_kaydet)
        
        self.setLayout(self.layout)
    
    def kaydet_ve_listele(self):
        ay_index = self.combo_ay.currentIndex()
        ay = self.combo_ay.itemText(ay_index)
        yil = int(self.combo_yil.currentText())

        try:
            gun_basi_km = float(self.input_gun_basi_km.text())
        except ValueError:
            QMessageBox.warning(self, "Hata", "Lütfen geçerli bir gün başı kilometre değeri girin.")
            return

        km_aralik_str = self.input_km_aralik.text().strip()
        try:
            parts = list(map(str.strip, km_aralik_str.split('-')))
            if len(parts) != 2:
                raise ValueError
            km_min, km_max = map(int, parts)
        except ValueError:
            QMessageBox.warning(self, "Hata", "Lütfen kilometre aralığını 'min-max' formatında girin (örn: 100-300).")
            return

        gorev_yeri = self.input_gorev_yeri.text()
        haftasonu_durumu = self.combo_haftasonu.currentText()

        workbook = Workbook()
        sheet = workbook.active
        
        sheet['A1'] = 'Tarih'
        sheet['B1'] = 'Gün Başı (km)'
        sheet['C1'] = 'Gün Sonu (km)'
        sheet['D1'] = 'Yapılan Kilometre'
        sheet['E1'] = 'Kontrol / İmza'
        sheet['F1'] = 'Göreve Gidilen Yer'
        
        first_day = datetime(yil, ay_index + 1, 1)
        last_day = datetime(yil, ay_index + 1, monthrange(yil, ay_index + 1)[1])
        current_day = first_day
        total_km = 0
        
        haftasonu_fill = PatternFill(start_color='FFC0C0', end_color='FFC0C0', fill_type='solid')
        
        while current_day <= last_day:
            gun_basi_tarih = current_day.strftime('%d/%m/%Y')
            
            if current_day.weekday() >= 5:  # Cumartesi veya Pazar
                if haftasonu_durumu == 'Çalışmıyor':
                    sheet.append([gun_basi_tarih, '', '', '', '', ''])
                    for col in range(1, 7):
                        sheet.cell(row=sheet.max_row, column=col).fill = haftasonu_fill

                else:
                    yapilan_km = random.randint(km_min, km_max)
                    gun_sonu_km = gun_basi_km + yapilan_km
                    total_km += yapilan_km
                    sheet.append([gun_basi_tarih, gun_basi_km, gun_sonu_km, yapilan_km, '', gorev_yeri])
                    gun_basi_km = gun_sonu_km
            else:
                yapilan_km = random.randint(km_min, km_max)
                gun_sonu_km = gun_basi_km + yapilan_km
                total_km += yapilan_km
                sheet.append([gun_basi_tarih, gun_basi_km, gun_sonu_km, yapilan_km, '', gorev_yeri])
                gun_basi_km = gun_sonu_km
                
            current_day += timedelta(days=1)
        
        # Toplam yapılan kilometreyi ekle
        total_row = ['', '', 'TOP. YAPILAN KM:', total_km]
        sheet.append(total_row)
        
        for col in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 20
            for cell in sheet[get_column_letter(col)]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        for row in range(2, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value:
                try:
                    gun_basi_tarih = datetime.strptime(sheet.cell(row=row, column=1).value, '%d/%m/%Y')
                    if gun_basi_tarih.weekday() >= 5:
                        for col in range(1, 7):
                            sheet.cell(row=row, column=col).fill = haftasonu_fill
                except ValueError:
                    continue

        # Kaydetme yeri ve adı seç
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Excel Dosyasını Kaydet", f"Servis_Raporu_{ay}_{yil}.xlsx", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        workbook.save(file_path)
        self.input_gun_basi_km.clear()
        self.input_km_aralik.clear()
        self.input_gorev_yeri.clear()

        QMessageBox.information(self, 'Bilgi', 'Excel dosyası başarıyla oluşturuldu!', QMessageBox.Ok)

# Uygulama çalıştırma kısmı
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelGenerator()
    window.show()
    sys.exit(app.exec_())
