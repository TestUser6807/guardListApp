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
  selectedMonth: Date | undefined = undefined;
  showModal = false;
  selectedPerson: PersonModel = new PersonModel(uuid.v4(), '', []);
  persons: PersonModel[] = []; 
  assignedDates: AssignedDateModel[] = [] //atananan tarihler ve kişiler
  dayWeight:number[]= [1.5,1, 1, 1, 0.75, 1.25, 2]; //pazardan cumartesiye gün ağırlıkları
  isEditing = false;
  dutyDayCountTolerance = 1; // Ortalama nöbet sayısına ek olarak izin verilen maksimum nöbet sayısı farkı  
  dutyDayWeightTolerance = 1; // Ortalama nöbet ağırlığına ek olarak izin verilen maksimum nöbet ağırlığı farkı 
  numberOfCycles = 5;
  // TODO use json or sqlite database 
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
      this.assignedDates = JSON.parse(storedAssignments).map((item: AssignedDateModel) => ({
        assignedDate: new Date(item.assignedDate),
        personName: item.personName,
        personId: item.personId
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
    this.selectedMonth = undefined;
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
        label: `${this.formatDate(day)} (${this.getWeekDay(day)})`, // Günün etiketini (tarih formatı) oluşturuyoruz
        value: new Date(day) // Günün tarihini değere atıyoruz
      });
    }
    localStorage.setItem('dutyDays', JSON.stringify(this.dutyDays));
    localStorage.removeItem('assignedDates');
    this.assignedDates = [];
    this.persons.forEach((p) => {
      p.notAvailableDays = [];
      p.dutyDays = [];
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));
    this.assignDates();
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
    this.resetDuty();
    this.assignDates();   
    this.closeModal();
  }
  resetDuty(){
    localStorage.removeItem('assignedDates');
    this.assignedDates = [];
    this.persons.forEach((p) => {
      p.dutyDays = [];
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));
  }
  removePerson(id: string) {
    const personIndex = this.persons.findIndex((p:PersonModel) => p.id === id);
    if (personIndex > -1) {
      this.persons.splice(personIndex, 1);
      localStorage.setItem('persons', JSON.stringify(this.persons));
      this.resetDuty();
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
  tryAssign(){
    const unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');
    this.assignShiftsToDaysWithoutAssignedShifts(unassignedDates);
  }
  resetAssignDates(){
    this.resetDuty();
    this.assignDates();
    this.tryAssign();
  }
  assignDates() {
      let personIndex = 0;

      this.sortPersonsByDutyDaysCountAndDutyWeight();
      this.sortDutyDaysByWeight();
      this.dutyDays.forEach((dutyDay) => {
        const person = this.persons[personIndex];
        if(this.userMustNotWorkTwoConsecutiveDays(person, dutyDay.value) &&
          this.userMustBeAvailable(person, dutyDay.value) &&
          this.noMoreThanTheAverageNumberOfDutysAreWorked(person) &&
          this.noMoreThanTheAverageNumberOfDutysWeightAreWorked(person) &&
          this.userMustNotWorkMoreThanOneShiftOnSameDay(person, dutyDay.value)
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
    
          person.dutyDays.push(dutyDay.value);
          localStorage.setItem('persons', JSON.stringify(this.persons));
          return;
        }

        // Person index'ini döngüsel şekilde güncelle
        const userHaveMoreNotAvailableDays = this.persons?.find((p) => {
            return p.notAvailableDays.length > this.dutyDays.length / 2;
          });

          if (userHaveMoreNotAvailableDays !== undefined &&
              this.userMustNotWorkTwoConsecutiveDays(person, dutyDay.value) &&
              this.userMustBeAvailable(person, dutyDay.value) &&
              this.noMoreThanTheAverageNumberOfDutysAreWorked(person) &&
              this.noMoreThanTheAverageNumberOfDutysWeightAreWorked(person)) {
            // Eğer nöbet tutabilecek günü az biri var ise ona tekrar nöbet verilir
            personIndex = this.persons.indexOf(userHaveMoreNotAvailableDays);
          } else {
            // Eğer personIndex, persons.length'ten büyükse sıfırlanır
            personIndex = (personIndex + 1) % this.persons.length; 
          }
      });

      // Nöbet tutulan tarihleri assignedDates'e set et
      this.setAssignedDates();

      // Atanamayan tarih varsa ata
     
      let unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');

      if(unassignedDates.length >0)
      {
        for(let i = 0; i<this.numberOfCycles; i++){
          unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');
          this.assignShiftsToDaysWithoutAssignedShifts(unassignedDates);
        }
      }
  }
  assignShiftsToDaysWithoutAssignedShifts(unassignedDates:AssignedDateModel[]){
    if(unassignedDates.length >0) {
       // kimseye atanmamış tarihleri sırala
      unassignedDates.sort((a,b)=>this.dayWeight[b.assignedDate.getDate()] - this.dayWeight[a.assignedDate.getDate()] );

      unassignedDates.forEach(dutyDay => {
        // Kişileri nöbet sayısı ve nöbet ağarlığına göre sırala
        this.sortPersonsByDutyDaysCountAndDutyWeight();
        // En düşük güne ve yüke sahip olan kişiyi al O kişiye dutyDay ekle
        const assignedPerson = this.persons[0];
        assignedPerson.dutyDays = assignedPerson.dutyDays || [];
        if(
          this.userMustNotWorkTwoConsecutiveDays(assignedPerson, dutyDay.assignedDate) &&
          this.userMustBeAvailable(assignedPerson, dutyDay.assignedDate) &&
          !assignedPerson.dutyDays.some(d => d.getTime() === dutyDay.assignedDate.getTime())
        )
          assignedPerson.dutyDays.push(dutyDay.assignedDate);
          this.setAssignedDates();
          localStorage.setItem('persons', JSON.stringify(this.persons));
      });
      
    }
  }
  noMoreThanTheAverageNumberOfDutysAreWorked(person: PersonModel): boolean {
    const averageCountPerPerson = this.dutyDays.length / this.persons.length;
    if(person.dutyDayCount + this.dutyDayCountTolerance > averageCountPerPerson)
      return false;

    return true;
  }
  noMoreThanTheAverageNumberOfDutysWeightAreWorked(person: PersonModel): boolean {
    // Tüm dutyDays'in ağırlıklarının toplamını hesapla
    const totalWeight = this.dutyDays.reduce((sum, dutyDay) => {
      const dayOfWeek = dutyDay.value.getDay();
      return sum + this.dayWeight[dayOfWeek];
    }, 0);
    console.log('Total Weight:', totalWeight);
    // Ortalama ağırlığı hesapla
    const averageWeightPerPerson = totalWeight / this.persons.length;

    // Person'un dutyDays'inin ağırlık toplamını hesapla
    const personWeight = person.dutyDays?.reduce((sum, dutyDate) => {
      const dayOfWeek = dutyDate.getDay();
      return sum + this.dayWeight[dayOfWeek];
    }, 0) || 0;

    // Eğer person'un ağırlığı, ortalamadan tolerans kadar büyükse, false döndür
    return (personWeight + this.dutyDayWeightTolerance  < averageWeightPerPerson);
  }
  userMustNotWorkTwoConsecutiveDays(person: PersonModel, date: Date): boolean {
      const previousDay = new Date(date);
      previousDay.setDate(date.getDate() - 1);

      const nextDay = new Date(date);
      nextDay.setDate(date.getDate() + 1);

      if(person.dutyDays){
        // Önceki ve sonraki günlerde nöbet tutmuş mu kontrol et
        const isPreviousDayWorked = person.dutyDays.some(dutyDate => 
            dutyDate.getFullYear() === previousDay.getFullYear() &&
            dutyDate.getMonth() === previousDay.getMonth() &&
            dutyDate.getDate() === previousDay.getDate()
        );

        const isNextDayWorked = person.dutyDays.some(dutyDate => 
            dutyDate.getFullYear() === nextDay.getFullYear() &&
            dutyDate.getMonth() === nextDay.getMonth() &&
            dutyDate.getDate() === nextDay.getDate()
        );
        // Eğer önceki veya sonraki gün çalıştıysa, ardışık nöbet engellenir
        if (isPreviousDayWorked || isNextDayWorked) {
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
  userMustNotWorkMoreThanOneShiftOnSameDay(person: PersonModel, date: Date): boolean {
    const dayOfWeek = date.getDay();
    const shiftsOnSameDay = person.dutyDays?.findIndex(dutyDate => dutyDate.getDay() === dayOfWeek) || [];
    return shiftsOnSameDay == -1;
  }
  sortPersonsByDutyDaysCountAndDutyWeight() {
    this.persons.sort((a, b) => {
      // İlk olarak dutyDayCount'a göre sıralama yap
      if (a.dutyDayCount !== b.dutyDayCount) {
        return a.dutyDayCount - b.dutyDayCount;
      }
      
      // Eğer dutyDayCount'lar eşitse, dutyDayWeight'e göre sıralama yap
      const aWeight = a.dutyDays?.reduce((sum, dutyDate) => sum + this.dayWeight[dutyDate.getDay()], 0) || 0;
      const bWeight = b.dutyDays?.reduce((sum, dutyDate) => sum + this.dayWeight[dutyDate.getDay()], 0) || 0;
      
      return aWeight - bWeight;  // Küçükten büyüğe sıralama
    });
  }
  sortDutyDaysByWeight(){
    this.dutyDays.sort((a,b)=>{
      return this.dayWeight[b.value.getDay()] - this.dayWeight[a.value.getDay()];  
    })
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
  get sortedDutyDays() {
    return this.dutyDays
      .slice() // orijinal diziyi kopyala (immutability)
      .sort((a, b) => a.value.getTime() - b.value.getTime());
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
