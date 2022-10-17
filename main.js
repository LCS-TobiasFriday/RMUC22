//Recommended reading to use these files.

/*

https://nodejs.dev/en/learn/
https://www.w3schools.com/nodejs/default.asp
https://code.visualstudio.com/docs/introvideos/basics

*/

//Load the modules that I would like to use

const fetch = require('axios'); //Package for Communicating With API. It is an HTTP Client

const exceljs = require('exceljs'); //Package for Working with Excel.


async function main(){

    //Enter your RM Login Credentials Below
    
    let credentials = {
        Username: 'Username',
        Password: 'Password',
        LocationID: 1
    };

    //Please put your company code below.

    const CompanyCode = 'CompanyCode'

    //Setup Options for Axios to communicate with the Rent Manager API
    
    let fetchTokenOptions = {
        url: `https://${CompanyCode}.api.rentmanager.com/Authentication/AuthorizeUser`,
        method: 'post',
        headers: {'Content-type':'application/json'},
        data: credentials
    };

    //Store returned value as variable

    let fetchTokenResponse = await fetch(fetchTokenOptions);

    //Extract the data/token from the returned object

    let TOKEN = fetchTokenResponse.data;

    //Print/Display/Log the output of the token

    console.log(fetchTokenResponse.data);

    //Pull Bill Information

    let fetchBillOptions = {

        url: `https://${CompanyCode}.api.rentmanager.com/Bills`,
        method: 'get',
        headers: {'Content-type':'application/json','X-RM12Api-ApiToken':TOKEN}

    };


    //Send the request to the API

    let fetchBillsResponse = await fetch(fetchBillOptions);

    let Bills = fetchBillsResponse.data;


    //Log off API

    let logoutOptions = {
        url: `https://${CompanyCode}.api.rentmanager.com/Authentication/DeAuthorize?token=${TOKEN}`,
        method: 'post',
        hheaders: {'Content-type':'application/json','X-RM12Api-ApiToken':TOKEN}
    };
    
    
    // Send the logout request

    await fetch(logoutOptions);

    //Extract the information from the Tenants

    let sheetRows = []; //Setup an array to store the rows for the sheet

    for (i in Bills){

        let Bill = Bills[i];

        let rowValues = {
            id: Bill.ID,
            account: Bill.AccountID,
            amount: Bill.Amount,
            transactiondate: Bill.TransactionDate,
            duedate: Bill.DueDate,
            accounttype: Bill.AccountType,
            createuser: Bill.CreateUserID,
            updateuser: Bill.UpdateUserID
        };

        sheetRows.push(rowValues);

    };



    //Access and setup the excel book

    let workbook = new exceljs.Workbook();

    let sheet = workbook.addWorksheet('Sheet1');

    sheet.columns = [
        {header: 'ID', key: 'id'},
        {header: 'Account', key: 'account'},
        {header: 'Amount', key:'amount'},
        {header: 'TransactionDate', key:'transactiondate'},
        {header: 'DueDate', key: 'duedate'},
        {header: 'VendorType', key: 'accounttype'},
        {header: 'CreateUser', key: 'createuser'},
        {header: 'UpdateUser', key: 'updateuser'}
    ]

    //Assign all of the tenant data to each row

    for (r in sheetRows){

        sheet.addRow(sheetRows[r])

    };

    //Enter the location on your computer you want to save the file to.

    const SaveLocation = 'C:/Program Files (x86)/BillsOutput.xlsx'

    //Save my work

    await workbook.xlsx.writeFile(SaveLocation);
    

    console.log('Done');

}


main();