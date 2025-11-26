import { CommonModule } from '@angular/common';
import { Component, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CalendarModule } from 'primeng/calendar';
import { MultiSelectModule } from 'primeng/multiselect';

import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { PersonModel } from './person.model';
import * as uuid from 'uuid';
import { AssignedDateModel } from './assigned-date.model';
import { ThisReceiver } from '@angular/compiler';

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
  showModal: boolean = false;
  selectedPerson: PersonModel = new PersonModel(uuid.v4(), '', []);
  persons: PersonModel[] = [];
  assignedDates: AssignedDateModel[] = [] //atananan tarihler ve kişiler
  dayWeight: number[] = [1.5, 1, 1, 1, 0.75, 1.25, 2]; //pazardan cumartesiye gün ağırlıkları
  isEditing: boolean = false;
  dutyDayCountTolerance: number = 1; // Ortalama nöbet sayısına ek olarak izin verilen maksimum nöbet sayısı farkı  
  dutyDayWeightTolerance: number = 0.75; // Ortalama nöbet ağırlığına ek olarak izin verilen maksimum nöbet ağırlığı farkı 
  numberOfCycles: number = 5;
  days: string[] = [
    'Pazar',
    'Pazartesi',
    'Salı',
    'Çarşamba',
    'Perşembe',
    'Cuma',
    'Cumartesi',
  ];
  assignedDatesView: (AssignedDateModel | null)[][] = Array.from({ length: 7 }, () => []);
  dayWeightAlert = signal('');
  dayCountAlert = signal('');

  // TODO use json or sqlite database 
  constructor() {
    const storedPersons = localStorage.getItem('persons');
    if (storedPersons && Array.isArray(storedPersons)) {
      this.persons = storedPersons;
    } else if (storedPersons) {
      try {
        const parsedPersons = JSON.parse(storedPersons).map((item: any) => new PersonModel(
          item.id,
          item.name,
          item.notAvailableDays ? item.notAvailableDays.map((d: string) => new Date(d)) : [],
          item.dutyDays ? item.dutyDays.map((d: string) => new Date(d)) : [],
          item.color
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
        personId: item.personId,
        color: item.color
      }));
    }

    const storedDutyDays = localStorage.getItem('dutyDays');
    if (storedDutyDays) {
      this.dutyDays = JSON.parse(storedDutyDays).map((item: any) => ({
        label: item.label,
        value: new Date(item.value),
      }));
    }
    this.setAssignedDatesView();
  }
  setAssignedDatesView() {
    // assignedDates boşsa işlem yapma
    if (!this.assignedDates || this.assignedDates.length === 0) {
      return; // Eğer assignedDates boşsa, fonksiyonu sonlandırabiliriz.
    }

    const firstAssignedDateDay = this.assignedDates[0]?.assignedDate.getDay(); // 0: Pazar, 1: Pazartesi, ..., 6: Cumartesi
    let emptySlots = 0;

    // İlk günün hangi günde olduğunu kontrol edip, boşluk sayısını belirliyoruz
    switch (firstAssignedDateDay) {
      case 0:
        emptySlots = 6; // Pazar ise, 6 boşluk
        break;
      case 1:
        emptySlots = 0; // Pazartesi ise, hiç boşluk yok
        break;
      case 2:
        emptySlots = 1; // Salı ise, 1 boşluk
        break;
      case 3:
        emptySlots = 2; // Çarşamba ise, 2 boşluk
        break;
      case 4:
        emptySlots = 3; // Perşembe ise, 3 boşluk (Mayıs 2025 örneği)
        break;
      case 5:
        emptySlots = 4; // Cuma ise, 4 boşluk
        break;
      case 6:
        emptySlots = 5; // Cumartesi ise, 5 boşluk
        break;
      default:
        emptySlots = 0; // Diğer günler için
        break;
    }

    // assignedDatesView'i 2 boyutlu dizi olarak oluşturuyoruz
    const totalRows = Math.ceil((this.assignedDates.length + emptySlots) / 7); // Kaç satır gerektiğini hesaplıyoruz (1 hafta = 7 gün)

    // assignedDatesView dizisini sıfırdan başlatıyoruz
    this.assignedDatesView = Array.from({ length: totalRows }, () => new Array(7).fill(null));
    
    let currentIndex = 0;

    // İlk satırı dolduruyoruz, boşlukları null ile dolduruyoruz
    for (let i = 0; i < 7; i++) {
      if (i < emptySlots) {
        this.assignedDatesView[0][i] = null; // İlk satırdaki boşlukları null ile dolduruyoruz
      } else {
        const assignedIndex = i - emptySlots;
        if (assignedIndex < this.assignedDates.length) {
          this.assignedDatesView[0][i] = this.assignedDates[currentIndex];
          currentIndex++;
        }
      }
    }

    // Diğer satırlarda, assignedDates dizisini yerleştiriyoruz
    for (let row = 1; row < totalRows; row++) {
      for (let col = 0; col < 7; col++) {
        if (currentIndex < this.assignedDates.length) {
          this.assignedDatesView[row][col] = this.assignedDates[currentIndex];
          currentIndex++;
        } else {
          this.assignedDatesView[row][col] = null; // Eğer assignedDates bittiğinde boşluk eklemek istiyorsanız
        }
      }
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
    this.assignedDatesView = Array.from({ length: 7 }, () => []);
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
    this.assignedDatesView = Array.from({ length: 7 }, () => []);
    this.persons.forEach((p) => {
      p.notAvailableDays = [];
      p.dutyDays = [];
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));
    this.assignDates();
  }
  openEditModal(person: PersonModel) {
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
  resetDuty() {
    localStorage.removeItem('assignedDates');
    this.assignedDates = [];
    this.assignedDatesView = Array.from({ length: 7 }, () => []);
    this.persons.forEach((p) => {
      p.dutyDays = [];
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));
  }
  removePerson(id: string) {
    const personIndex = this.persons.findIndex((p: PersonModel) => p.id === id);
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
    this.assignedDatesView = Array.from({ length: 7 }, () => []);
    localStorage.removeItem('assignedDates');
  }
  resetAllPersonNotAvailableDays() {
    this.persons.forEach((p) => {
      p.notAvailableDays = []
    });
    localStorage.setItem('persons', JSON.stringify(this.persons));
    this.assignDates();
  }
  tryAssign() {
    for (let i = 0; i < 5; i++) {
      this.sortPersonsByDutyDaysCountAndDutyWeight();
      // Son kişiyi al (sonuncu kişi)
      const lastPerson = this.persons[this.persons.length - 1];

      // Eğer dutyDays varsa
      if (lastPerson.dutyDays && lastPerson.dutyDays.length > 0) {
        // Her dutyDay için dayWeight hesapla ve en yüksek olanı bul
        let maxWeight = -Infinity;  // Başlangıç için çok düşük bir değer
        let maxWeightIndex = -1;  // En yüksek ağırlığa sahip olan günü bulmak için index

        lastPerson.dutyDays.forEach((dutyDay, index) => {
          const dayWeight = this.dayWeight[dutyDay.getDay()];  // `getDay()` günün haftadaki numarasını döndürür
          if (dayWeight > maxWeight) {
            maxWeight = dayWeight;
            maxWeightIndex = index;
          }
        });

        // En yüksek ağırlığa sahip günü sil
        if (maxWeightIndex !== -1) {
          lastPerson.dutyDays.splice(maxWeightIndex, 1);  // `splice` ile o günü çıkar
        }
      }
    }
    this.setAssignedDates();
    let unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');
    this.assignShiftsToDaysWithoutAssignedShifts(unassignedDates)
  }
  hardAssign() {
    let unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');
    this.assignShiftsToDaysWithoutAssignedShifts(unassignedDates, true)    
  }
  resetAssignDates() {
    this.resetDuty();
    this.assignDates();
  }
  assignDates() {
    let personIndex = 0;

    this.sortPersonDesc();
    this.sortDutyDays();
    console.log(this.dutyDays)
    this.dutyDays.forEach((dutyDay) => {
      const person = this.persons[personIndex];
      if (this.userMustNotWorkTwoConsecutiveDays(person, dutyDay.value) &&
        this.userMustBeAvailable(person, dutyDay.value) &&
        this.noMoreThanTheAverageNumberOfDutysAreWorked(person) &&
        this.noMoreThanTheAverageNumberOfDutysWeightAreWorked(person) &&
        this.userMustNotWorkMoreThanOneShiftOnSameDay(person, dutyDay.value) &&
        this.userOnDutyOnThursdayMustNotBeOnDutyOnTheWeekend(person, dutyDay.value)
      ) {
        person.dutyDays = person.dutyDays || [];
        person.dutyDays.push(dutyDay.value);
        localStorage.setItem('persons', JSON.stringify(this.persons));
        return;
      } else {
        personIndex = (personIndex + 1) % this.persons.length;
      }
    });

    // Nöbet tutulan tarihleri assignedDates'e set et
    this.setAssignedDates();

    // Atanamayan tarih varsa atamaya çalış
    let unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');
    this.assignShiftsToDaysWithoutAssignedShifts(unassignedDates);
    // Daha iyi atama
    this.doBetterDutyWeightAssign();
  }
  doBetterDutyWeightAssign() {
    // 1) En fazla dutyDays (nöbet) olan kullanıcıyı bulma (maxUser)
    const maxWeightedDutyUser = this.persons.reduce((maxUser, currentUser) =>
      !maxUser || currentUser.dutyDayWeight > maxUser.dutyDayWeight
        ? currentUser
        : maxUser
    );

    // 2) En az dutyDays (nöbet) olan kullanıcıyı bulma (minUser)
    const minWeightedDutyUser = this.persons.reduce((minUser, currentUser) =>
      !minUser || currentUser.dutyDayWeight < minUser.dutyDayWeight
        ? currentUser
        : minUser
    );

    if (!maxWeightedDutyUser.dutyDays?.length || !minWeightedDutyUser.dutyDays?.length) {
      return;
    }

    const maxDays = maxWeightedDutyUser.dutyDays;
    const minDays = minWeightedDutyUser.dutyDays;

    // 3) Swap yapmak için uygun kombinasyonları dene
    let swapped = false;

    // Kombinasyonları sırayla dene: ilk başta maxUser'ın en fazla nöbeti ile minUser'ın en az nöbeti
    const combinations = [];

    // Max günlerin sırasıyla denenecek swap kombinasyonları
    for (let i = 0; i < maxDays.length; i++) {
      for (let j = 0; j < minDays.length; j++) {
        combinations.push([i, j]);  // [maxUserın indexi, minUserın indexi]
      }
    }

    // 4) Swap kombinasyonlarını sırayla dene
    for (const [maxIndex, minIndex] of combinations) {
      const maxUserMaxDay = maxDays[maxIndex]; // maxUser'ın ilgili en fazla nöbet günü
      const minUserMinDay = minDays[minIndex]; // minUser'ın ilgili en az nöbet günü

      // Max ve Min günlerinin ağırlıklarını karşılaştır
      const maxDayWeight = this.dayWeight[maxUserMaxDay.getDay()]; // maxUser'ın gününün ağırlığı
      const minDayWeight = this.dayWeight[minUserMinDay.getDay()]; // minUser'ın gününün ağırlığı

      // MinUser'ın günü, MaxUser'ın gününden daha düşük ağırlığa sahip olmalı
      if (minDayWeight >= maxDayWeight) {
        continue; // Eğer minUser'ın gününün ağırlığı, maxUser'ınkinden fazla veya eşitse, swap yapma
      }

      // Swap için gerekli kurallar
      const maxUserValid =
        this.userMustNotWorkTwoConsecutiveDays(maxWeightedDutyUser, minUserMinDay) &&
        this.userMustBeAvailable(maxWeightedDutyUser, minUserMinDay);

      const minUserValid =
        this.userMustNotWorkTwoConsecutiveDays(minWeightedDutyUser, maxUserMaxDay) &&
        this.userMustBeAvailable(minWeightedDutyUser, maxUserMaxDay);

      // Eğer geçerli swap yapılabilirse
      if (maxUserValid && minUserValid) {
        // Yer değiştir
        const maxDayIndex = maxDays.indexOf(maxUserMaxDay);
        const minDayIndex = minDays.indexOf(minUserMinDay);

        if (maxDayIndex !== -1 && minDayIndex !== -1) {
          // Swap işlemi
          const temp = maxDays[maxDayIndex];
          maxDays[maxDayIndex] = minDays[minDayIndex];
          minDays[minDayIndex] = temp;

          swapped = true;  // Swap başarılı
          break;  // Swap başarıyla yapıldığında döngüden çık
        }
      }
    }

    // Kaydet
    localStorage.setItem("persons", JSON.stringify(this.persons));
    this.setAssignedDates();
  }
  assignShiftsToDaysWithoutAssignedShifts(unassignedDates: AssignedDateModel[], assignHard?: boolean) {
    if (unassignedDates.length > 0) {
      assignHard = assignHard == null ? false : true;
      // Unassigned tarihleri sırala
      unassignedDates.sort(
        (a, b) =>
          this.dayWeight[b.assignedDate.getDate()] -
          this.dayWeight[a.assignedDate.getDate()]
      );

      // Her tarih için bir kez atama yapılacak
      for (const dutyDay of unassignedDates) {

        // Kişileri nöbet sayılarına göre sırala
        this.sortPersonsByDutyDaysCountAndDutyWeight();

        for (const person of this.persons) {
          person.dutyDays = person.dutyDays || [];
          if (
            this.userMustNotWorkTwoConsecutiveDays(person, dutyDay.assignedDate) &&
            this.userMustBeAvailable(person, dutyDay.assignedDate) &&
            (assignHard ? true :  this.noMoreThanTheAverageNumberOfDutysAreWorked(person)) &&
            (assignHard ? true : this.userOnDutyOnThursdayMustNotBeOnDutyOnTheWeekend(person, dutyDay.assignedDate)) &&
            (assignHard ? true : this.userMustNotWorkMoreThanOneShiftOnSameDay(person, dutyDay.assignedDate)) &&
            !person.dutyDays.some(d => d.getTime() === dutyDay.assignedDate.getTime())
          ) {
            // Atamayı yap
            person.dutyDays.push(dutyDay.assignedDate);
            this.setAssignedDates();
            break;
          }
        }
      }

      // Kaydet
      localStorage.setItem('persons', JSON.stringify(this.persons));
    }
  }
  noMoreThanTheAverageNumberOfDutysAreWorked(person: PersonModel): boolean {
    const averageCountPerPerson = this.dutyDays.length / this.persons.length;
    if (person.dutyDayCount + this.dutyDayCountTolerance > averageCountPerPerson)
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
    return (personWeight + this.dutyDayWeightTolerance < averageWeightPerPerson);
  }
  userMustNotWorkTwoConsecutiveDays(person: PersonModel, date: Date): boolean {
    const previousDay = new Date(date);
    previousDay.setDate(date.getDate() - 1);

    const nextDay = new Date(date);
    nextDay.setDate(date.getDate() + 1);

    if (person.dutyDays) {
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
  userMustBeAvailable(person: PersonModel, date: Date): boolean {
    if (person.notAvailableDays) {
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
    //ignore pazartesi salı çarşamba
    if (dayOfWeek == 1 || dayOfWeek == 2 || dayOfWeek == 3)
      return true;

    const shiftsOnSameDay = person.dutyDays?.findIndex(dutyDate => dutyDate.getDay() === dayOfWeek) || [];
    return shiftsOnSameDay == -1;
  }
  userOnDutyOnThursdayMustNotBeOnDutyOnTheWeekend(person: PersonModel, date: Date): boolean {
    const dayOfWeek = date.getDay();
    const previousThursday = new Date(date);

    // Eğer bugün Perşembe ise
    if (dayOfWeek === 4) {
      const saturday = new Date(date);
      saturday.setDate(date.getDate() + 2);  // 2 gün sonrası Cumartesi

      const sunday = new Date(date);
      sunday.setDate(date.getDate() + 3);  // 3 gün sonrası Pazar

      // Cumartesi veya Pazar günü için nöbet verilmişse, false dön
      if (person.dutyDays) {
        const isSaturdayOnDuty = person.dutyDays.some(dutyDate =>
          dutyDate.getFullYear() === saturday.getFullYear() &&
          dutyDate.getMonth() === saturday.getMonth() &&
          dutyDate.getDate() === saturday.getDate()
        );

        const isSundayOnDuty = person.dutyDays.some(dutyDate =>
          dutyDate.getFullYear() === sunday.getFullYear() &&
          dutyDate.getMonth() === sunday.getMonth() &&
          dutyDate.getDate() === sunday.getDate()
        );

        if (isSaturdayOnDuty || isSundayOnDuty) {
          return false;  // Eğer Cumartesi veya Pazar nöbeti varsa, false dön
        }
      }
    }

    // Eğer tarih Cumartesi veya Pazar ise, önceki Perşembe'yi bul
    if (dayOfWeek === 6) { // Cumartesi
      previousThursday.setDate(date.getDate() - 2); // Perşembe'ye 2 gün ekle
    } else if (dayOfWeek === 0) { // Pazar
      previousThursday.setDate(date.getDate() - 3); // Perşembe'ye 3 gün ekle
    } else {
      return true; // Diğer günlerde bir şey yapmaya gerek yok
    }

    // Perşembe günü çalışıldığını kontrol et
    if (person.dutyDays) {
      const isPreviousThursdayWorked = person.dutyDays.some(dutyDate =>
        dutyDate.getFullYear() === previousThursday.getFullYear() &&
        dutyDate.getMonth() === previousThursday.getMonth() &&
        dutyDate.getDate() === previousThursday.getDate()
      );

      if (isPreviousThursdayWorked) {
        return false; // Eğer Perşembe günü çalışıldıysa, hafta sonu çalışılmamalı
      }
    }

    return true; // Eğer Perşembe günü çalışılmadıysa, hafta sonu çalışabilir
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
  sortPersonDesc() {
    this.persons.sort((a, b) => {
      // İlk olarak notAvailableDays.length'e göre sıralama yap
      if (a.notAvailableDays.length !== b.notAvailableDays.length) {
        return b.notAvailableDays.length - a.notAvailableDays.length;  // En fazla notAvailableDays olan önce gelsin
      }

      // Eğer notAvailableDays'lar eşitse, dutyDayCount'a göre sıralama yap
      if (a.dutyDayCount !== b.dutyDayCount) {
        return a.dutyDayCount - b.dutyDayCount;
      }

      // Eğer dutyDayCount'lar eşitse, dutyDayWeight'e göre sıralama yap
      const aWeight = a.dutyDays?.reduce((sum, dutyDate) => sum + this.dayWeight[dutyDate.getDay()], 0) || 0;
      const bWeight = b.dutyDays?.reduce((sum, dutyDate) => sum + this.dayWeight[dutyDate.getDay()], 0) || 0;

      return aWeight - bWeight;  // Küçükten büyüğe sıralama
    });
  }
  sortDutyDays() {
    // Günlerin sıralama ağırlıkları (Cumartesi = 0, Perşembe = 1, Pazar = 2, Cuma = 3, diğer günler ise sonraki sırada)
    const dayWeight:{ [key: number]: number } = {
      6: 0, // Cumartesi
      4: 1, // Perşembe
      0: 2, // Pazar
      5: 3, // Cuma
      1: 4, // Pazartesi
      2: 5, // Salı
      3: 6, // Çarşamba
    };

    this.dutyDays.sort((a, b) => {
      // `a.value.getDay()` ve `b.value.getDay()` ile hafta içindeki günü alıyoruz
      const aDay = a.value.getDay();
      const bDay = b.value.getDay();
      
      // Öncelikle, günlerin ağırlıklarına göre sıralama yapıyoruz
      return dayWeight[aDay] - dayWeight[bDay];
    });
  }
  setAssignedDates() {
    this.assignedDates = []; // Öncelikle assignedDates dizisini sıfırlıyoruz.
    this.assignedDatesView = Array.from({ length: 7 }, () => []);

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
            dutyDate,
            person.color
          );
        } else {
          // Eğer tarih yoksa, yeni atama ekliyoruz
          this.assignedDates.push(new AssignedDateModel(
            person.id,
            person.name,
            dutyDate,
            person.color
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
    this.setAssignedDatesView();
    this.alert();
  }
  getWeekDay(date: Date): string {
    return this.days[date.getDay()];
  }
  alert(){
    const maxWeightedUser = this.persons.reduce((maxUser, currentUser) => {
      return currentUser.dutyDayWeight > maxUser.dutyDayWeight ? currentUser : maxUser;
    }, this.persons[0]);
     const minWeightedUser = this.persons.reduce((minUser, currentUser) => {
      return currentUser.dutyDayWeight < minUser.dutyDayWeight ? currentUser : minUser;
    }, this.persons[0]);
    let minUserAvaliableDayCount = this.dutyDays.length - minWeightedUser.notAvailableDays.length;
    let userCanDuty = (minUserAvaliableDayCount / 2) -1 >= minWeightedUser.dutyDayCount
    if(maxWeightedUser.dutyDayWeight - minWeightedUser.dutyDayWeight > 2 && userCanDuty){
      this.dayWeightAlert.set(`${maxWeightedUser.name} ${maxWeightedUser.dutyDayWeight} ağarlığında nöbet tutarken  ${minWeightedUser.name} ${minWeightedUser.dutyDayWeight} ağarlığında nöbet tutuyor düzeltme gerekebilir!`);
    }else{
      this.dayWeightAlert.set('');
    }

    const maxCountUser = this.persons.reduce((maxUser, currentUser) => {
      return currentUser.dutyDayCount > maxUser.dutyDayCount ? currentUser : maxUser;
    }, this.persons[0]);
     const minCountUser = this.persons.reduce((minUser, currentUser) => {
      return currentUser.dutyDayCount < minUser.dutyDayCount ? currentUser : minUser;
    }, this.persons[0]);
    minUserAvaliableDayCount = this.dutyDays.length - minCountUser.notAvailableDays.length;
    userCanDuty = (minUserAvaliableDayCount / 2) -1 >= minCountUser.dutyDayCount
    if(maxCountUser.dutyDayCount - minCountUser.dutyDayCount > 2 && userCanDuty){
      this.dayCountAlert.set(`${maxWeightedUser.name} ${maxWeightedUser.dutyDayCount} nöbet tutarken  ${minWeightedUser.name} ${minWeightedUser.dutyDayCount} nöbet tutuyor düzeltme gerekebilir!`);
    }else{
      this.dayCountAlert.set('');
    }
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
  get sortedPersons() {
    return this.persons
      .slice() // orijinal diziyi kopyala (immutability)
      .sort((a, b) => a.name.localeCompare(b.name)); // Alfabetik sıralama
  }
  get unassignedDatesCount() {
    const unassignedDates = this.assignedDates.filter(d => d.personId == 'Kimse atanmadı');
    return unassignedDates.length;
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
