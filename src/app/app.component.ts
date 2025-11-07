import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CalendarModule } from 'primeng/calendar';
import { MultiSelectModule } from 'primeng/multiselect';

import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { PersonModel } from './person.model';
import * as uuid from 'uuid';
import { AssignedDateModel } from './assigned-date.model';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, CalendarModule, MultiSelectModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  dutyDays: { label: string; value: Date }[] = [];
  selectedMonth: Date = new Date();
  showModal = false;
  selectedPerson: PersonModel = new PersonModel(uuid.v4(), '', []);
  persons: PersonModel[] = []; 
  assignedDates: AssignedDateModel[] = [] //atananan tarihler ve kişiler
  dayWeight:number[]= [1.5,1, 1, 1, 0.75, 1.25, 2]; //pazardan cumartesiye gün ağırlıkları
  isEditing = false;
  dutyDayCountTolerance = 1; // Ortalama nöbet sayısına ek olarak izin verilen maksimum nöbet sayısı farkı  
  dutyDayWeightTolerance = 0.5; // Ortalama nöbet ağırlığına ek olarak izin verilen maksimum nöbet ağırlığı farkı 
  // TODO use json database 
  constructor() {
    const storedPersons = localStorage.getItem('persons');
    if (storedPersons && Array.isArray(storedPersons)) {
    this.persons = storedPersons; 
    } else if(storedPersons) {
      try {
        const parsedPersons = JSON.parse(storedPersons).map((item: any) => new PersonModel(
          item.id,
          item.name,
          item.notAvailableDays ? item.notAvailableDays.map((d: string) => new Date(d)) : [],
          item.dutyDays ? item.dutyDays.map((d: string) => new Date(d)) : []
        ));
        if (Array.isArray(parsedPersons)) {
          this.persons = parsedPersons;
        }
      } catch (e) {
        console.error('JSON parse hatası:', e);
      }
    }


    const storedAssignments = localStorage.getItem('assignedDates');
    if (storedAssignments) {
      this.assignedDates = JSON.parse(storedAssignments).map((item: any) => ({
        date: new Date(item.date),
        person: item.person,
      }));
    }

    const storedDutyDays = localStorage.getItem('dutyDays');
    if (storedDutyDays) {
      this.dutyDays = JSON.parse(storedDutyDays).map((item: any) => ({
        label: item.label,
        value: new Date(item.value),
      }));
    }
  }
  reset() {
    localStorage.removeItem('assignedDates');
    localStorage.removeItem('dutyDays');
    this.dutyDays = [];
    this.assignedDates = [];
    this.selectedPerson = new PersonModel(uuid.v4(), '', []);
    this.persons.forEach((p) => {
      p.dutyDays = [];
      p.notAvailableDays = [];
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));

  }
  onDutyDaySelected(event: Date) {
    const selectedMonth = event.getMonth(); // Seçilen ay (0-11 arası)
    const selectedYear = event.getFullYear(); // Seçilen yıl

    // Ayın başı ve sonunu belirliyoruz
    const firstDayOfMonth = new Date(selectedYear, selectedMonth, 1);
    const lastDayOfMonth = new Date(selectedYear, selectedMonth + 1, 0);

    // Ayın ilk gününden son gününe kadar döngü
    this.dutyDays = [];
    for (let day = firstDayOfMonth; day <= lastDayOfMonth; day.setDate(day.getDate() + 1)) {
      this.dutyDays.push({
        label: this.formatDate(day), // Günün etiketini (tarih formatı) oluşturuyoruz
        value: new Date(day) // Günün tarihini değere atıyoruz
      });
    }
    localStorage.setItem('dutyDays', JSON.stringify(this.dutyDays));
    this.resetAllPersonNotAvailableDays();
  }
  openEditModal(person:PersonModel) {
    this.isEditing = true;
    this.showModal = true;
    this.selectedPerson = new PersonModel(
      person.id,
      person.name,
      person.notAvailableDays,
      person.dutyDays
    );
  }
  openCreateModel() { 
    this.isEditing = false; 
    this.showModal = true;
    this.selectedPerson = new PersonModel(uuid.v4(), '', []);
  }
  closeModal() {
    this.showModal = false;
  }
  savePerson() {
    const person = new PersonModel(
      this.isEditing === false
        ? this.selectedPerson.id
        : uuid.v4(),
      this.selectedPerson.name,
      this.selectedPerson.notAvailableDays,
      this.selectedPerson.dutyDays
    );

    if (this.isEditing) {
      const index = this.persons.findIndex(p => p.id == this.selectedPerson.id);
      if (index !== -1) {
        this.persons[index] = person;
      }
    } else {
      this.persons.push(person);
    }
    this.selectedPerson = person;
    localStorage.setItem('persons', JSON.stringify(this.persons));

    this.assignDates();
    this.closeModal();
  }
  removePerson(id: string) {
    const personIndex = this.persons.findIndex((p:PersonModel) => p.id === id);
    if (personIndex > -1) {
      this.persons.splice(personIndex, 1);
      localStorage.setItem('persons', JSON.stringify(this.persons));
      this.assignDates();
    }
  }
  removeAllPerson() {
    this.persons = [];
    localStorage.setItem('persons', JSON.stringify(this.persons));
    this.assignedDates = [];
    localStorage.removeItem('assignedDates');
  }
  resetAllPersonNotAvailableDays() {
    this.persons.forEach((p) => {
      p.notAvailableDays = []
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));
    this.assignDates();
  }
  assignDates() {
    this.dutyDays.forEach((dutyDay) => {
      this.persons.forEach((person) => {
        if(this.userMustNotWorkTwoConsecutiveDays(person, dutyDay.value) &&
          this.userMustBeAvailable(person, dutyDay.value) &&
          this.noMoreThanTheAverageNumberOfDutysAreWorked(person) &&
          this.noMoreThanTheAverageNumberOfDutysWeightAreWorked(person) 
        ) {
          person.dutyDays = person.dutyDays || [];
          // Bu tarih daha önce atanmış mı kontrol et, atanmışsa eskiyi sil
          if (this.assignedDates.some(ad => ad.assignedDate.getTime() === dutyDay.value.getTime())) {
              this.persons.forEach((p) => {
                  // person.dutyDays dizisinde bu tarihe sahip olup olmadığını kontrol et
                  const pddIndex = p.dutyDays?.findIndex(pdd => pdd.getTime() === dutyDay.value.getTime());

                  // Eğer tarih bulunduysa, dutyDays'ten sil
                  if (pddIndex !== undefined && pddIndex >= 0) {
                      p.dutyDays?.splice(pddIndex, 1);
                  }
              });
          }
          // Todo gelişti  Atama yap
          person.dutyDays.push(dutyDay.value);
          localStorage.setItem('persons', JSON.stringify(this.persons));
          return;
        }
      });
    });
    this.setAssignedDates();
  }

  noMoreThanTheAverageNumberOfDutysAreWorked(person: PersonModel): boolean {
    const averageCountPerPerson = this.dutyDays.length / this.persons.length;

    if(person.dutyDayCount > averageCountPerPerson + this.dutyDayCountTolerance)
      return false;

    return true;
  }
  noMoreThanTheAverageNumberOfDutysWeightAreWorked(person: PersonModel): boolean {
    // Tüm dutyDays'in ağırlıklarının toplamını hesapla
    const totalWeight = this.dutyDays.reduce((sum, dutyDay) => {
      const dayOfWeek = dutyDay.value.getDay();
      return sum + this.dayWeight[dayOfWeek];
    }, 0);

    // Ortalama ağırlığı hesapla
    const averageWeightPerPerson = totalWeight / this.persons.length;

    // Person'un dutyDays'inin ağırlık toplamını hesapla
    const personWeight = person.dutyDays?.reduce((sum, dutyDate) => {
      const dayOfWeek = dutyDate.getDay();
      return sum + this.dayWeight[dayOfWeek];
    }, 0) || 0;

    // Eğer person'un ağırlığı, ortalamadan tolerans kadar büyükse, false döndür
    return (personWeight < (averageWeightPerPerson + this.dutyDayWeightTolerance));
  }


  userMustNotWorkTwoConsecutiveDays(person: PersonModel, date: Date): boolean {
    const previousDay = new Date(date);
    previousDay.setDate(date.getDate() - 1);

    if (person.dutyDays) {
      const isPreviousDayWorked = person.dutyDays.some(dutyDate => {
        return (
          dutyDate.getFullYear() === previousDay.getFullYear() &&
          dutyDate.getMonth() === previousDay.getMonth() &&
          dutyDate.getDate() === previousDay.getDate()
        );
      });

      if (isPreviousDayWorked) {
        return false;
      }
    }

    return true;
  }
  userMustBeAvailable(person:PersonModel, date:Date): boolean {
    if(person.notAvailableDays)
    {
      for (let i = 0; i < person.notAvailableDays.length; i++) {
        const naDate = person.notAvailableDays[i];
        if (
          naDate.getFullYear() === date.getFullYear() &&
          naDate.getMonth() === date.getMonth() &&
          naDate.getDate() === date.getDate()
        ) {
          return false;
        }
      }
    }
    return true;
  }

  setAssignedDates() {
    this.assignedDates = []; // Öncelikle assignedDates dizisini sıfırlıyoruz.

    // Person'lar üzerinden geçiyoruz ve her bir person için dutyDays'leri kontrol ediyoruz
    this.persons.forEach((person) => {
      person.dutyDays?.forEach((dutyDate) => {
        // Tarih daha önce assignedDates dizisinde var mı diye kontrol ediyoruz
        const assignedDateIndex = this.assignedDates.findIndex(ad => ad.assignedDate.getTime() === dutyDate.getTime());
        
        if (assignedDateIndex !== -1) {
          // Eğer tarih zaten varsa, person'ı güncelliyoruz (yani güncellenmiş atama yapıyoruz)
          this.assignedDates[assignedDateIndex] = new AssignedDateModel(
            person.id,
            person.name,
            dutyDate
          );
        } else {
          // Eğer tarih yoksa, yeni atama ekliyoruz
          this.assignedDates.push(new AssignedDateModel(
            person.id,
            person.name,
            dutyDate
          ));
        }
      });
    });

    // dutyDays dizisindeki tüm tarihler için "Kimse atanmadı" ataması yapılmış mı diye kontrol ediyoruz
    this.dutyDays.forEach((dutyDay) => {
      const isAssigned = this.assignedDates.some(ad => ad.assignedDate.getTime() === dutyDay.value.getTime());
      if (!isAssigned) {
        // Eğer o tarih için kimse atanmadıysa, "Kimse atanmadı" olarak yeni öğe ekliyoruz
        this.assignedDates.push(new AssignedDateModel(
          'Kimse atanmadı',
          'Kimse atanmadı',
          dutyDay.value
        ));
      }
    });

    // Tarihe göre sırala
    this.assignedDates.sort((a, b) => a.assignedDate.getTime() - b.assignedDate.getTime());

    // assignedDates dizisini localStorage'a kaydediyoruz
    localStorage.setItem('assignedDates', JSON.stringify(this.assignedDates));
  }
  getWeekDay(date: Date): string {
    const days = [
      'Pazar',
      'Pazartesi',
      'Salı',
      'Çarşamba',
      'Perşembe',
      'Cuma',
      'Cumartesi',
    ];
    return days[date.getDay()];
  }
  formatDate(date?: Date): string {
    if (!date) return '';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
  }
  exportExcel(): void {
    // 1. Veri hazırlanıyor
    const worksheetData = this.assignedDates.map((item) => ({
      Tarih: this.formatDate(item.assignedDate),
      Atanan: item.personName,
    }));

    // 2. En eski ve en yeni tarih
    const timestamps = this.assignedDates.map((item) =>
      new Date(item.assignedDate).getTime()
    );
    const minDate = this.formatDate(new Date(Math.min(...timestamps)));
    const maxDate = this.formatDate(new Date(Math.max(...timestamps)));
    const title = `${minDate} - ${maxDate} Nöbet Listesi`;

    // 3. Workbook ve Worksheet oluştur
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Nöbet Listesi');

    // 4. Başlık ekle (A1 ve B1 hücrelerini birleştir)
    worksheet.mergeCells('A1:B1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = title;
    titleCell.font = { bold: true, size: 16 };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // 5. Sütun başlıklarını ekle (2. satır)
    worksheet.getRow(2).values = ['Tarih', 'Atanan'];
    worksheet.getRow(2).font = { bold: true };

    // 6. Kişi isimlerine renk ata
    const uniquePeople = [
      ...new Set(
        this.assignedDates.map((item) => item.personName.trim().toLowerCase())
      ),
    ];
    const colorPalette = [
      'FFB6C1',
      'ADD8E6',
      '90EE90',
      'FFFF99',
      'FFD700',
      'FFA07A',
      'DDA0DD',
      '00CED1',
      'F08080',
      'E0FFFF',
      'C0C0C0',
      'FFC0CB',
      '98FB98',
      'AFEEEE',
      'FFFACD',
      '0046FF',
      '73C8D2',
      'F5F1DC',
      'FF9013',
      '004030',
      '4A9782',
      'DCD0A8',
      'FFF9E5',
    ];
    const personColors: { [key: string]: string } = {};
    uniquePeople.forEach((person, i) => {
      personColors[person] = colorPalette[i % colorPalette.length];
    });

    // 7. Verileri ekle ve satırları renklendir
    worksheetData.forEach((item, index) => {
      const rowIndex = index + 3; // veri 3. satırdan başlıyor
      const row = worksheet.getRow(rowIndex);

      const dateParts = item.Tarih.split('.'); // ["01", "10", "2025"]
      const tarihAsDate = new Date(
        +dateParts[2],
        +dateParts[1] - 1,
        +dateParts[0]
      );
      const weekDay = this.getWeekDay(tarihAsDate);

      row.getCell(1).value = `${item.Tarih} (${weekDay})`;
      row.getCell(2).value = item.Atanan;

      const personKey = item.Atanan.trim().toLowerCase();
      const fillColor = personColors[personKey];

      // Satırı renklendir
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: fillColor },
        };
        cell.font = {
          color: { argb: 'FF000000' }, // Siyah yazı
        };
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };
      });

      row.commit();
    });

    // 8. Sütun genişliklerini ayarla
    worksheet.columns = [
      { key: 'Tarih', width: 15 },
      { key: 'Atanan', width: 25 },
    ];

    // 9. Dosyayı oluştur ve kaydet
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      saveAs(blob, 'Nobet_Listesi.xlsx');
    });
  }
}
