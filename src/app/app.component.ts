import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CalendarModule } from 'primeng/calendar';
import { MultiSelectModule } from 'primeng/multiselect';

import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';


@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, CalendarModule, MultiSelectModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  rawDates: Date[] = [];
  dates: { label: string; value: Date }[] = [];
  selectedMonth: Date = new Date();
  selectedPersonDates: { label: string; value: Date }[] = [];
  selectAll = true;
  showModal = false;

  personName: string = '';
  persons: { name: string; dates: Date[] }[] = [];
  assignedDates: { date: Date; person: string }[] = [];

  dayWeight:number[]= [1.5,1, 1, 1, 0.75, 1.25, 2]; //pazardan cumartesiye
  isDayWeightedActive : boolean = true;
  // Güncelleme modları
  isEditing = false;
  editingIndex: number | null = null;

  exportExcel(): void {
    // 1. Veri hazırlanıyor
    const worksheetData = this.assignedDates.map((item) => ({
      Tarih: this.formatDate(item.date),
      Atanan: item.person,
    }));

    // 2. En eski ve en yeni tarih
    const timestamps = this.assignedDates.map((item) =>
      new Date(item.date).getTime()
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
        this.assignedDates.map((item) => item.person.trim().toLowerCase())
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

  constructor() {
    const storedPersons = localStorage.getItem('persons');
    if (storedPersons) {
      this.persons = JSON.parse(storedPersons).map((p: any) => ({
        name: p.name,
        dates: p.dates.map((d: string) => new Date(d)),
      }));
    }

    const storedAssignments = localStorage.getItem('assignedDates');
    if (storedAssignments) {
      this.assignedDates = JSON.parse(storedAssignments).map((item: any) => ({
        date: new Date(item.date),
        person: item.person,
      }));
    }
  }

  removePerson(person: { name: string; dates: Date[] }) {
    const index = this.persons.findIndex((p) => p.name === person.name);
    if (index > -1) {
      this.persons.splice(index, 1);
      localStorage.setItem('persons', JSON.stringify(this.persons));
      this.assignDates();
    }
  }

  reset() {
    localStorage.removeItem('assignedDates');
    this.rawDates = [];
    this.dates = [];
    this.selectedPersonDates = [];
    this.assignedDates = [];
    this.isDayWeightedActive = true;
  }

  onDatesSelected(newDates: Date[]) {
    this.rawDates = newDates.sort((a, b) => a.getTime() - b.getTime()); // Tarihleri sırala
    this.dates = this.rawDates.map((d) => ({
      label: `${this.formatDate(d)} (${this.getWeekDay(d)})`, // label'a gün ekle
      value: d,
    }));

    this.assignDates();
  }
onMonthSelected(event: any) {
  console.log('Selected Month:', event);
  // You can perform further logic with the selected date here
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

  getDutyCount(name: string): number {
    return this.assignedDates.filter((item) => item.person === name).length;
  }
  getWeightedCount(name: string): number {
    let totalWeight = 0;
    this.assignedDates.forEach((item) => {  
      if (item.person === name) {
        const dayOfWeek = item.date.getDay();
        totalWeight += this.dayWeight[dayOfWeek];
      }
    });
    return totalWeight;
  }

normalizeDate(date: Date): Date {
  if (date instanceof Date && !isNaN(date.getTime())) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }
  return new Date(NaN);
}


  formatDate(date: Date): string {
    return new Intl.DateTimeFormat('tr-TR').format(this.normalizeDate(date));
  }
updatePersonDates(): void {
    // Normalize ve geçersizleri temizle
    const selectedDates: Date[] = this.dates
      .map((d) => this.normalizeDate(d.value))
      .filter((dt) => dt instanceof Date && !isNaN(dt.getTime()));

    // Eğer hiç geçerli tarih yoksa hiçbir değişiklik yapma
    if (selectedDates.length === 0) return;

    // Her kişiye ayrı kopya ver (referans paylaşımı olmasın)
    this.persons = this.persons.map((person) => ({
      ...person,
      dates: selectedDates.map((d) => new Date(d.getTime())),
    }));

    // localStorage'a ISO string olarak yaz (constructor ile uyumlu parsing için)
    const toStore = this.persons.map((p) => ({
      name: p.name,
      dates: p.dates.map((d) => d.toISOString()),
    }));
    localStorage.setItem('persons', JSON.stringify(toStore));

    this.assignDates();
  }

  openModal(index?: number) {
    if (typeof index === 'number') {
      const person = this.persons[index];
      this.personName = person.name;
      // dates içinden, person's dates'ine uygun label+value objelerini seç
      this.selectedPersonDates = this.dates.filter((d) =>
        person.dates.some(
          (pd) =>
            this.normalizeDate(pd).toDateString() ===
            this.normalizeDate(d.value).toDateString()
        )
      );

      this.isEditing = true;
      this.editingIndex = index;
    } else {
      this.personName = '';
      this.selectedPersonDates = [...this.dates]; // Tüm tarihler seçili
      this.isEditing = false;
      this.editingIndex = null;
    }
    this.showModal = true;
  }

  closeModal() {
    this.showModal = false;
    this.isEditing = false;
    this.editingIndex = null;
  }

  toggleSelectAll(event: any) {
    this.selectAll = event.target.checked;
    this.selectedPersonDates = this.selectAll ? [...this.dates] : [];
  }

  savePerson() {
    if (!this.personName || this.selectedPersonDates.length === 0) return;

    const person = {
      name: this.personName,
      dates: this.selectedPersonDates.map((d) => this.normalizeDate(d.value)),
    };

    if (this.isEditing && this.editingIndex !== null) {
      this.persons[this.editingIndex] = person;
    } else {
      this.persons.push(person);
    }

    localStorage.setItem('persons', JSON.stringify(this.persons));

    this.assignDates();
    this.closeModal();
  }

  // assignDates() {
  //   const assignments: { date: Date; person: string }[] = [];
  //   const personCounts: { [key: string]: number } = {};
  //   const sortedDates = this.rawDates
  //     .slice()
  //     .sort((a, b) => a.getTime() - b.getTime());

  //   // Her kişi için başta nöbet sayısı 0
  //   this.persons.forEach((p) => {
  //     personCounts[p.name] = 0;
  //   });

  //   for (let i = 0; i < sortedDates.length; i++) {
  //     const date = sortedDates[i];

  //     // Bu tarihi tutabilecek kişiler
  //     const available = this.persons.filter((p) =>
  //       p.dates.some((d) => new Date(d).toDateString() === date.toDateString())
  //     );

  //     // Önceki gün atanmış kişi varsa onu çıkar
  //     const previousAssignment = assignments[i - 1];
  //     let filtered = available;

  //     if (previousAssignment) {
  //       filtered = available.filter(
  //         (p) => p.name !== previousAssignment.person
  //       );
  //     }

  //     if (filtered.length > 0) {
  //       // En az nöbet sayısına sahip olanları bul
  //       const minCount = Math.min(...filtered.map((p) => personCounts[p.name]));
  //       const candidates = filtered.filter(
  //         (p) => personCounts[p.name] === minCount
  //       );

  //       // Rastgele seç
  //       const chosen =
  //         candidates[Math.floor(Math.random() * candidates.length)];

  //       personCounts[chosen.name]++;
  //       assignments.push({ date, person: chosen.name });
  //     } else {
  //       // Kimse atanamıyorsa
  //       assignments.push({ date, person: 'Kimse atanmadı' });
  //     }
  //   }

  //   this.assignedDates = assignments;
  //   localStorage.setItem('assignedDates', JSON.stringify(assignments));
  // }
  //TODO fixed
assignDates() {
  const MAX_ATTEMPTS = 500;

  const normalizeDateStr = (d: Date) =>
    this.normalizeDate(new Date(d)).toDateString();

  const makeAttempt = () => {
    const assignments: { date: Date; person: string }[] = [];
    const personCounts: { [key: string]: number } = {};
    const weightCounts: { [key: string]: number } = {};
    const personWeekdayAssigned: { [person: string]: { [day: number]: number } } = {};

    const totalDays = this.rawDates.length;
    const numPersons = this.persons.length;
    if (numPersons === 0 || totalDays === 0) {
      this.assignedDates = [];
      localStorage.setItem('assignedDates', JSON.stringify(this.assignedDates));
      return { success: true, assignments };
    }

    const sortedDates = this.rawDates.slice().sort((a, b) => a.getTime() - b.getTime());

    // Başlangıç sayaçları
    this.persons.forEach((p) => {
      personCounts[p.name] = 0;
      weightCounts[p.name] = 0;
      personWeekdayAssigned[p.name] = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
    });

    // Ortalamalar
    const totalWeight = sortedDates.reduce((s, dt) => s + this.dayWeight[dt.getDay()], 0);
    const avgDutyCount = totalDays / numPersons;
    const avgWeightedCount = totalWeight / numPersons;

    const allowedMinDuty = Math.max(0, Math.floor(avgDutyCount) - 1);
    const allowedMaxDuty = Math.ceil(avgDutyCount) + 1;
    const allowedMinWeighted = avgWeightedCount - 0.5;
    const allowedMaxWeighted = avgWeightedCount + 0.5;

    // Gün başına tarih listesi ve toplam sayıları
    const dayDates: { [day: number]: Date[] } = { 0: [], 1: [], 2: [], 3: [], 4: [], 5: [], 6: [] };
    sortedDates.forEach((d) => dayDates[d.getDay()].push(d));

    // Her gün için kişiler arasında hedef dağılımı (eşit ve rastgele)
    const targetsByDay: { [day: number]: { [person: string]: number } } = {};
    for (let day = 0; day < 7; day++) {
      const count = dayDates[day].length;
      const base = Math.floor(count / numPersons);
      let remainder = count % numPersons;

      const shuffled = this.persons.map((p) => p.name).sort(() => Math.random() - 0.5);
      targetsByDay[day] = {};
      this.persons.forEach((p) => (targetsByDay[day][p.name] = base));
      for (let r = 0; r < remainder; r++) {
        targetsByDay[day][shuffled[r]]++;
      }
    }

    // Atama döngüsü: tarihler sırayla atanıyor
    for (let i = 0; i < sortedDates.length; i++) {
      const date = sortedDates[i];
      const day = date.getDay();
      const dayWeightValue = this.dayWeight[day];
      const prev = assignments[i - 1];

      // Eligible: o tarihte müsait olan kişiler
      let eligible = this.persons.filter((p) =>
        p.dates.some((pd) => normalizeDateStr(pd) === normalizeDateStr(date))
      );

      // Önceki gün atanmış kişiyi çıkar
      if (prev) {
        eligible = eligible.filter((p) => p.name !== prev.person);
      }

      // Eğer eligible boşsa direkt 'Kimse atanmadı'
      if (eligible.length === 0) {
        assignments.push({ date, person: 'Kimse atanmadı' });
        continue;
      }

      // Mevcut en düşük nöbet sayısını al; böylece hiçbir kişi diğerlerinden 2 fazla olamaz
      const currentMinCount = Math.min(...Object.values(personCounts));
      // Bir kişinin alabileceği en yüksek nöbet sayısı (hem avg sınırı hem de diğerlerine göre fark)
      const dynamicMaxDuty = Math.min(allowedMaxDuty, currentMinCount + 1);

      // Filtreleme aşamaları: önce tüm kısıtlarla, sonra kademeli gevşet
      const tryFilters = [
        // 1) tam kısıtlar: weekday hedefi aşmasın, duty <= dynamicMaxDuty, weighted <= allowedMaxWeighted
        (pname: string) =>
          personWeekdayAssigned[pname][day] < targetsByDay[day][pname] &&
          personCounts[pname] + 1 <= dynamicMaxDuty &&
          weightCounts[pname] + dayWeightValue <= allowedMaxWeighted,
        // 2) izin verilen weekday hedefini gevşet (sadece duty ve weighted kontrol)
        (pname: string) =>
          personCounts[pname] + 1 <= dynamicMaxDuty &&
          weightCounts[pname] + dayWeightValue <= allowedMaxWeighted,
        // 3) sadece duty limit kontrolü (farkı koruyacak şekilde)
        (pname: string) => personCounts[pname] + 1 <= dynamicMaxDuty,
        // 4) sadece weighted limit kontrolü
        (pname: string) => weightCounts[pname] + dayWeightValue <= allowedMaxWeighted,
        // 5) hiçbir limit (ama önceki gün kontrolü halen sağlanıyor)
        (_pname: string) => true,
      ];

      let chosenName: string | null = null;
      for (const filter of tryFilters) {
        const candidates = eligible
          .map((p) => p.name)
          .filter((pn) => filter(pn));
        if (candidates.length === 0) continue;

        // En az assignment/weighted olanları öne al, sonra rastgele seç
        candidates.sort((a, b) => {
          const va = personCounts[a] + weightCounts[a];
          const vb = personCounts[b] + weightCounts[b];
          if (va !== vb) return va - vb;
          return Math.random() - 0.5;
        });

        chosenName = candidates[0];
        break;
      }

      if (!chosenName) {
        assignments.push({ date, person: 'Kimse atanmadı' });
        continue;
      }

      // Atamayı kaydet
      personCounts[chosenName]++;
      weightCounts[chosenName] += dayWeightValue;
      personWeekdayAssigned[chosenName][day]++;
      assignments.push({ date, person: chosenName });
    }

    // Son kontrol: herkes izin verilen aralıkta mı?
    let ok = true;
    const counts = this.persons.map(p => personCounts[p.name]);
    const maxCount = Math.max(...counts);
    const minCount = Math.min(...counts);

    // 1) duty sınırları ve aralarındaki fark (biri 4 diğer 6 olamaz)
    for (const p of this.persons) {
      const c = personCounts[p.name];
      if (c < allowedMinDuty || c > allowedMaxDuty) {
        ok = false;
        break;
      }
    }
    if (ok && (maxCount - minCount) > 1) ok = false;

    // 2) weighted sınırları
    if (ok) {
      for (const p of this.persons) {
        const w = weightCounts[p.name];
        if (w < allowedMinWeighted - 1e-9 || w > allowedMaxWeighted + 1e-9) {
          ok = false;
          break;
        }
      }
    }

    // 3) ardışık gün kontrolü (Kimse atanmadı hariç)
    if (ok) {
      for (let i = 1; i < assignments.length; i++) {
        const prev = assignments[i - 1];
        const cur = assignments[i];
        if (prev.person !== 'Kimse atanmadı' && cur.person !== 'Kimse atanmadı' && prev.person === cur.person) {
          ok = false;
          break;
        }
      }
    }

    return { success: ok, assignments, personCounts, weightCounts };
  };

  // Tekrarlamalı denemeler
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const result = makeAttempt();
    if (result.success) {
      this.assignedDates = result.assignments;
      localStorage.setItem('assignedDates', JSON.stringify(this.assignedDates));
      console.log(`assignDates succeeded on attempt ${attempt}`);
      return;
    }
  }

  // Eğer tüm denemeler başarısızsa, en son denemeyi kabul et (kullanıcıya uyarı)
  const last = makeAttempt();
  this.assignedDates = last.assignments;
  localStorage.setItem('assignedDates', JSON.stringify(this.assignedDates));
  console.warn('assignDates: constraints could not be fully satisfied within attempt limit. Final assignments saved.');
}
}
