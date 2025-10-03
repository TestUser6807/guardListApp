import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CalendarModule } from 'primeng/calendar';
import { MultiSelectModule } from 'primeng/multiselect';
import * as XLSX from 'xlsx';
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
  selectedPersonDates: { label: string; value: Date }[] = [];
  selectAll = true;
  showModal = false;

  personName: string = '';
  persons: { name: string; dates: Date[] }[] = [];
  assignedDates: { date: Date; person: string }[] = [];

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

    // 3. Worksheet oluştur ve veriyi A2 hücresinden başlat
    const worksheet = XLSX.utils.aoa_to_sheet([]);
    XLSX.utils.sheet_add_json(worksheet, worksheetData, {
      origin: 'A2',
      skipHeader: false,
    });

    // 4. Başlığı A1 ve B1 hücresinde birleştirerek yaz
    XLSX.utils.sheet_add_aoa(worksheet, [[title]], { origin: 'A1' });
    (worksheet as any)['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];

    (worksheet as any)['A1'].s = {
      font: { bold: true, sz: 18 },
      alignment: { horizontal: 'center' },
    };

    // 5. Her kişiye renk ata (normalize ederek: trim ve lowercase)
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

    // İsimleri normalize et (trim ve lowercase)
    const uniquePeople = [
      ...new Set(
        this.assignedDates.map((item) => item.person.trim().toLowerCase())
      ),
    ];

    const personColors: { [key: string]: string } = {};
    uniquePeople.forEach((person, i) => {
      personColors[person] = colorPalette[i % colorPalette.length];
    });

    // 6. Satırları renklendir
    const startRow = 2; // veri A2’den başlıyor, yani Excel'de 2. satır
    this.assignedDates.forEach((item, index) => {
      const row = startRow + index;
      const personKey = item.person.trim().toLowerCase();
      const fillColor = personColors[personKey].replace('#', '');

      ['A', 'B'].forEach((col) => {
        const cellRef = `${col}${row}`;
        if (worksheet[cellRef]) {
          (worksheet[cellRef] as any).s = {
            fill: {
              patternType: 'solid',
              fgColor: { rgb: fillColor },
            },
            font: {
              color: { rgb: '000000' },
            },
            alignment: {
              vertical: 'center',
              horizontal: 'left',
              wrapText: true,
            },
          };
        }
      });
    });

    // 7. Workbook oluştur
    const workbook: XLSX.WorkBook = {
      SheetNames: ['Nöbet Listesi'],
      Sheets: { 'Nöbet Listesi': worksheet },
    };

    // 8. Excel dosyasını oluştur
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
      cellStyles: true,
    });

    // 9. Dosyayı kaydet
    this.saveAsExcelFile(excelBuffer, 'Nobet_Listesi');
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    saveAs(data, fileName + '.xlsx');
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
  }

  onDatesSelected(newDates: Date[]) {
    this.rawDates = newDates.sort((a, b) => a.getTime() - b.getTime()); // Tarihleri sırala
    this.dates = this.rawDates.map((d) => ({
      label: `${this.formatDate(d)} (${this.getWeekDay(d)})`, // label'a gün ekle
      value: d,
    }));
    this.assignDates();
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

  normalizeDate(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  formatDate(date: Date): string {
    return new Intl.DateTimeFormat('tr-TR').format(this.normalizeDate(date));
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

  assignDates() {
    const assignments: { date: Date; person: string }[] = [];
    const personCounts: { [key: string]: number } = {};
    const sortedDates = this.rawDates
      .slice()
      .sort((a, b) => a.getTime() - b.getTime());

    // Her kişi için başta nöbet sayısı 0
    this.persons.forEach((p) => {
      personCounts[p.name] = 0;
    });

    for (let i = 0; i < sortedDates.length; i++) {
      const date = sortedDates[i];

      // Bu tarihi tutabilecek kişiler
      const available = this.persons.filter((p) =>
        p.dates.some((d) => new Date(d).toDateString() === date.toDateString())
      );

      // Önceki gün atanmış kişi varsa onu çıkar
      const previousAssignment = assignments[i - 1];
      let filtered = available;

      if (previousAssignment) {
        filtered = available.filter(
          (p) => p.name !== previousAssignment.person
        );
      }

      if (filtered.length > 0) {
        // En az nöbet sayısına sahip olanları bul
        const minCount = Math.min(...filtered.map((p) => personCounts[p.name]));
        const candidates = filtered.filter(
          (p) => personCounts[p.name] === minCount
        );

        // Rastgele seç
        const chosen =
          candidates[Math.floor(Math.random() * candidates.length)];

        personCounts[chosen.name]++;
        assignments.push({ date, person: chosen.name });
      } else {
        // Kimse atanamıyorsa
        assignments.push({ date, person: 'Kimse atanmadı' });
      }
    }

    this.assignedDates = assignments;
    localStorage.setItem('assignedDates', JSON.stringify(assignments));
  }
}
