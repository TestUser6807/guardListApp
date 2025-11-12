export class PersonModel {
    id: string = '';
    name: string = '';
    notAvailableDays: Date[] = [];
    dutyDays?: Date[] = [];
    
    constructor(id: string, name: string, notAvailableDays?: Date[], dutyDays?: Date[]) {
        this.id = id;
        this.name = name;
        this.notAvailableDays = notAvailableDays || [];
        this.dutyDays = dutyDays || [];
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
