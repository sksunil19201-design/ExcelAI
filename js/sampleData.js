/*
=================================================
 Excel AI - Sample Data Library
 Version : 1.0
 Author  : Sunil Kumar
=================================================
*/

const SAMPLE_DATA = {

    /*=================================================
      SALES DATASET
    =================================================*/

    sales: [
        { Region: "North", Product: "Laptop", SalesPerson: "John", Month: "Jan", Quantity: 12, Amount: 1200 },
        { Region: "South", Product: "Mouse", SalesPerson: "David", Month: "Jan", Quantity: 20, Amount: 800 },
        { Region: "East", Product: "Keyboard", SalesPerson: "Alex", Month: "Jan", Quantity: 15, Amount: 950 },
        { Region: "West", Product: "Monitor", SalesPerson: "Sara", Month: "Jan", Quantity: 8, Amount: 1500 },
        { Region: "North", Product: "Printer", SalesPerson: "Emma", Month: "Feb", Quantity: 5, Amount: 700 },
        { Region: "South", Product: "Laptop", SalesPerson: "Mike", Month: "Feb", Quantity: 10, Amount: 1800 },
        { Region: "East", Product: "Mouse", SalesPerson: "Sophia", Month: "Feb", Quantity: 25, Amount: 650 },
        { Region: "West", Product: "Keyboard", SalesPerson: "James", Month: "Mar", Quantity: 14, Amount: 1050 },
        { Region: "North", Product: "Monitor", SalesPerson: "Robert", Month: "Mar", Quantity: 6, Amount: 2200 },
        { Region: "South", Product: "Printer", SalesPerson: "William", Month: "Mar", Quantity: 4, Amount: 900 }
    ],

    /*=================================================
      EMPLOYEE DATASET
    =================================================*/

    employees: [
        { ID: 101, Name: "John", Department: "HR", Salary: 50000, City: "Delhi" },
        { ID: 102, Name: "David", Department: "Finance", Salary: 65000, City: "Mumbai" },
        { ID: 103, Name: "Alex", Department: "IT", Salary: 72000, City: "Pune" },
        { ID: 104, Name: "Sara", Department: "Sales", Salary: 58000, City: "Hyderabad" },
        { ID: 105, Name: "Emma", Department: "IT", Salary: 81000, City: "Bangalore" },
        { ID: 106, Name: "James", Department: "Finance", Salary: 55000, City: "Delhi" },
        { ID: 107, Name: "Sophia", Department: "HR", Salary: 60000, City: "Chennai" }
    ],

    /*=================================================
      ATTENDANCE DATASET
    =================================================*/

    attendance: [
        { Employee: "John", Date: "01-Jan", Status: "Present" },
        { Employee: "David", Date: "01-Jan", Status: "Absent" },
        { Employee: "Alex", Date: "01-Jan", Status: "Present" },
        { Employee: "Sara", Date: "01-Jan", Status: "Leave" },
        { Employee: "Emma", Date: "01-Jan", Status: "Present" },
        { Employee: "James", Date: "01-Jan", Status: "Present" },
        { Employee: "Sophia", Date: "01-Jan", Status: "Absent" }
    ],

    /*=================================================
      INVENTORY DATASET
    =================================================*/

    inventory: [
        { Item: "Laptop", Category: "Electronics", Stock: 25, Price: 800 },
        { Item: "Mouse", Category: "Accessories", Stock: 120, Price: 20 },
        { Item: "Keyboard", Category: "Accessories", Stock: 80, Price: 35 },
        { Item: "Monitor", Category: "Electronics", Stock: 40, Price: 220 },
        { Item: "Printer", Category: "Office", Stock: 15, Price: 180 }
    ],

    /*=================================================
      STUDENT DATASET
    =================================================*/

    students: [
        { RollNo: 1, Name: "Rahul", Maths: 85, Science: 72, English: 81 },
        { RollNo: 2, Name: "Priya", Maths: 91, Science: 88, English: 94 },
        { RollNo: 3, Name: "Aman", Maths: 70, Science: 76, English: 73 },
        { RollNo: 4, Name: "Neha", Maths: 89, Science: 91, English: 95 },
        { RollNo: 5, Name: "Rohan", Maths: 77, Science: 74, English: 80 }
    ],

    /*=================================================
      BANKING DATASET
    =================================================*/

    banking: [
        { Account: "SB1001", Customer: "John", Debit: 500, Credit: 1000, Balance: 5000 },
        { Account: "SB1002", Customer: "David", Debit: 800, Credit: 1200, Balance: 4200 },
        { Account: "SB1003", Customer: "Alex", Debit: 300, Credit: 700, Balance: 6100 },
        { Account: "SB1004", Customer: "Sara", Debit: 1000, Credit: 2000, Balance: 7500 },
        { Account: "SB1005", Customer: "Emma", Debit: 600, Credit: 900, Balance: 3900 }
    ]

};


/*=================================================
  Return Dataset
=================================================*/

function getSampleData(datasetName) {

    if (SAMPLE_DATA.hasOwnProperty(datasetName)) {
        return SAMPLE_DATA[datasetName];
    }

    console.warn("Dataset not found :", datasetName);

    return [];

}


/*=================================================
  Get Dataset Names
=================================================*/

function getDatasetNames() {

    return Object.keys(SAMPLE_DATA);

}


/*=================================================
  Clone Dataset
=================================================*/

function cloneDataset(datasetName) {

    return JSON.parse(JSON.stringify(getSampleData(datasetName)));

}


console.log("✅ sampleData.js Loaded Successfully");
