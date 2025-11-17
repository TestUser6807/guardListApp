export class PersonModel {
    id: string = '';
    name: string = '';
    notAvailableDays: Date[] = [];
    dutyDays?: Date[] = [];
    color:string='';
    private colorPalette = [
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
    private static usedColors: Set<string> = new Set();
    constructor(id: string, name: string, notAvailableDays?: Date[], dutyDays?: Date[],color?:string) {
        this.id = id;
        this.name = name;
        this.notAvailableDays = notAvailableDays || [];
        this.dutyDays = dutyDays || [];
        if (color) {
                this.color = color;
            } else {
                // Renk paletinden benzersiz bir renk seç
                this.color = this.getUniqueColor();
            }
    }

    private getUniqueColor(): string {
        let newColor: string;
        
        // Kullanıcıya daha önce atanmış olan renkleri geçmemek için
        do {
            newColor = '#' + this.colorPalette[Math.floor(Math.random() * this.colorPalette.length)];
        } while (PersonModel.usedColors.has(newColor));

        // Yeni rengi kullanılmış renkler listesine ekle
        PersonModel.usedColors.add(newColor);
        
        return newColor;
    }

    // Duty day count'ı getter olarak döndürüyoruz.
    get dutyDayCount(): number {
        return this.dutyDays?.length || 0;
    }

    // Duty day weight'i getter olarak döndürüyoruz.
    get dutyDayWeight(): number {
        return this.calculateDutyDayWeight();
    }

    // Ağırlık hesaplama fonksiyonu
    public calculateDutyDayWeight(): number {
        // Günlerin ağırlıklarını belirliyoruz: [Pazar, Pazartesi, Salı, Çarşamba, Perşembe, Cuma, Cumartesi]
        const dayWeight: number[] = [1.5, 1, 1, 1, 0.75, 1.25, 2];

        // Eğer dutyDays varsa, her bir günü değerlendiriyoruz
        if (this.dutyDays && this.dutyDays.length > 0) {
            return this.dutyDays.reduce((totalWeight, date) => {
                const dayOfWeek = date.getDay(); // 0: Pazar, 1: Pazartesi, ..., 6: Cumartesi
                const dayWeightValue = dayWeight[dayOfWeek]; // O günün ağırlığını alıyoruz
                return totalWeight + dayWeightValue; // O günün ağırlığını toplam ağırlığa ekliyoruz
            }, 0); // Başlangıç değeri 0
        }

        // Eğer dutyDays boşsa, ağırlık 0 olarak döner
        return 0;
    }
    public calculateDutyWeekDayCount(dayIndex: number): number {
        if (!this.dutyDays || this.dutyDays.length === 0)
            return 0;
        return this.dutyDays.filter(dutyDate => dutyDate.getDay() === dayIndex).length;
    }

}
