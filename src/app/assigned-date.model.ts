export class AssignedDateModel{
    personId: string = '';
    personName: string = '';
    assignedDate: Date = new Date();

    constructor(personId: string, personName:string, assignedDate: Date){
        this.personId = personId;
        this.personName = personName;
        this.assignedDate = assignedDate;
    }
}