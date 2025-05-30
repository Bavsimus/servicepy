import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QComboBox, QMessageBox, QFileDialog, QSpinBox, QFormLayout, QGroupBox
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
        self.setWindowTitle('Çoklu Araç için Excel Tablosu (Tek Sayfa)')
        self.setGeometry(100, 100, 800, 600)
        self.layout = QVBoxLayout()

        self.combo_yil = QComboBox()
        self.combo_yil.addItems([str(yil) for yil in range(2020, 2031)])

        self.combo_ay = QComboBox()
        self.combo_ay.addItems([
            'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran',
            'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'
        ])

        self.spin_arac_sayisi = QSpinBox()
        self.spin_arac_sayisi.setRange(1, 10)
        self.spin_arac_sayisi.valueChanged.connect(self.arac_formlarini_guncelle)

        self.layout.addWidget(QLabel("Yıl Seçin:"))
        self.layout.addWidget(self.combo_yil)
        self.layout.addWidget(QLabel("Ay Seçin:"))
        self.layout.addWidget(self.combo_ay)
        self.layout.addWidget(QLabel("Araç Sayısı:"))
        self.layout.addWidget(self.spin_arac_sayisi)

        self.arac_formlari_container = QVBoxLayout()
        self.arac_formlari_group = QGroupBox("Araç Bilgileri")
        self.arac_formlari_group.setLayout(self.arac_formlari_container)
        self.layout.addWidget(self.arac_formlari_group)

        self.button_kaydet = QPushButton("Excel'i Oluştur")
        self.button_kaydet.clicked.connect(self.kaydet_ve_listele)
        self.layout.addWidget(self.button_kaydet)

        self.setLayout(self.layout)
        self.arac_inputlar = []
        self.arac_formlarini_guncelle()

    def arac_formlarini_guncelle(self):
        for i in reversed(range(self.arac_formlari_container.count())):
            self.arac_formlari_container.itemAt(i).widget().setParent(None)
        self.arac_inputlar.clear()

        for i in range(self.spin_arac_sayisi.value()):
            grup = QGroupBox(f"Araç {i + 1}")
            form = QFormLayout()

            input_baslangic_km = QLineEdit()
            input_km_aralik = QLineEdit()
            input_gorev_yeri = QLineEdit()
            input_haftasonu = QComboBox()
            input_haftasonu.addItems(['Çalışıyor', 'Çalışmıyor'])

            form.addRow("Gün Başı Kilometre:", input_baslangic_km)
            form.addRow("KM Aralığı (örn: 90-100):", input_km_aralik)
            form.addRow("Görev Yeri:", input_gorev_yeri)
            form.addRow("Hafta Sonu Durumu:", input_haftasonu)

            grup.setLayout(form)
            self.arac_formlari_container.addWidget(grup)

            self.arac_inputlar.append({
                "baslangic_km": input_baslangic_km,
                "km_aralik": input_km_aralik,
                "gorev_yeri": input_gorev_yeri,
                "haftasonu": input_haftasonu
            })

    def kaydet_ve_listele(self):
        ay_index = self.combo_ay.currentIndex()
        yil = int(self.combo_yil.currentText())
        gun_sayisi = monthrange(yil, ay_index + 1)[1]

        wb = Workbook()
        ws = wb.active
        ws.title = "Tüm Araçlar"

        row_cursor = 1  # Başlangıç satırı
        fill = PatternFill(start_color='FFC0C0', end_color='FFC0C0', fill_type='solid')

        for index, arac in enumerate(self.arac_inputlar):
            try:
                km_baslangic = float(arac["baslangic_km"].text())
                km_min, km_max = map(int, arac["km_aralik"].text().split('-'))
                gorev_yeri = arac["gorev_yeri"].text().strip()
                haftasonu_durumu = arac["haftasonu"].currentText()
            except Exception:
                QMessageBox.warning(self, "Hata", f"Araç {index+1} için veriler eksik ya da hatalı.")
                return

            ws.cell(row=row_cursor, column=1).value = f"Araç {index + 1}"
            row_cursor += 1

            headers = ['Tarih', 'Gün Başı (km)', 'Gün Sonu (km)', 'Yapılan KM', 'İmza', 'Görev Yeri']
            for col, header in enumerate(headers, 1):
                ws.cell(row=row_cursor, column=col, value=header)
                ws.cell(row=row_cursor, column=col).alignment = Alignment(horizontal='center', vertical='center')
            row_cursor += 1

            current_km = km_baslangic
            total_km = 0

            for day in range(1, gun_sayisi + 1):
                tarih = datetime(yil, ay_index + 1, day)
                str_tarih = tarih.strftime('%d/%m/%Y')

                if tarih.weekday() >= 5:
                    if haftasonu_durumu == 'Çalışmıyor':
                        # Boş ve boyalı satır
                        for col in range(1, 7):
                            ws.cell(row=row_cursor, column=col).fill = fill
                        ws.cell(row=row_cursor, column=1, value=str_tarih)
                        row_cursor += 1
                        continue
                    else:
                        # Dolu ama yine de boyalı satır
                        yapilan_km = random.randint(km_min, km_max)
                        gun_sonu_km = current_km + yapilan_km
                        total_km += yapilan_km

                        ws.append([str_tarih, current_km, gun_sonu_km, yapilan_km, '', gorev_yeri])
                        for col in range(1, 7):
                            ws.cell(row=row_cursor, column=col).fill = fill
                        row_cursor += 1
                        current_km = gun_sonu_km
                        continue
                yapilan_km = random.randint(km_min, km_max)
                gun_sonu_km = current_km + yapilan_km
                total_km += yapilan_km

                ws.append([str_tarih, current_km, gun_sonu_km, yapilan_km, '', gorev_yeri])
                row_cursor += 1
                current_km = gun_sonu_km

            ws.append(['', '', 'TOPLAM KM:', total_km])
            row_cursor += 3  # 2 satır boşluk + başlık için 1

        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Excel Dosyasını Kaydet", f"Tum_Araclar_{yil}_{ay_index + 1}.xlsx", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        wb.save(file_path)
        QMessageBox.information(self, "Başarılı", "Excel dosyası başarıyla kaydedildi!")

# Ana uygulama
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelGenerator()
    window.show()
    sys.exit(app.exec_())
