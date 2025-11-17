export class AssignedDateModel{
    personId: string = '';
    personName: string = '';
    assignedDate: Date = new Date();
    color:string = '#ff0000';

    constructor(personId: string, personName:string, assignedDate: Date,color?:string){
        this.personId = personId;
        this.personName = personName;
        this.assignedDate = assignedDate;
        this.color = color ?? '#ff0000' ;
    }
}