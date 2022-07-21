const {Builder, By, Key, until} = require('selenium-webdriver');

const ExcelJS = require('exceljs');

const sheet = './Book1.xlsx';
const wb = new ExcelJS.Workbook();



// let username = wb.xlsx.readFile('./Book1.xlsx').then(() => { 
//     console.dir(wb.worksheets[0].getCell('B2'). value)
// })
let password;
let username;
let currentLocation;
let legalStatus;
let educationSector;
let lastName;
let firstName;
let gender;
let dateOfBirth;
let passportNumber;
let passportCountry;
let passportNationality;
let passportIssuanceDate;
let passportExpiryDate;
let passportIssuer;
let birthCity;
let birthState;
let birthCountry;
let relationshipStatus;
let immigrationOffice;

wb.xlsx.readFile('./Book1.xlsx').then(() => { 
    username = wb.worksheets[0].getCell('B2').value;
    password = wb.worksheets[0].getCell('B3').value;
    currentLocation = wb.worksheets[0].getCell('B4').value;
    legalStatus = wb.worksheets[0].getCell('B5').value;
    educationSector = wb.worksheets[0].getCell('B6').value;
    lastName = wb.worksheets[0].getCell('B7').value;
    firstName = wb.worksheets[0].getCell('B8').value;
    gender = wb.worksheets[0].getCell('B9').value;
    dateOfBirth = wb.worksheets[0].getCell('B10').value;
    passportNumber = wb.worksheets[0].getCell('B11').value;
    passportCountry = wb.worksheets[0].getCell('B12').value;
    passportNationality = wb.worksheets[0].getCell('B13').value;
    passportIssuanceDate = wb.worksheets[0].getCell('B14').value;
    passportExpiryDate = wb.worksheets[0].getCell('B15').value;
    passportIssuer = wb.worksheets[0].getCell('B16').value;
    birthCity = wb.worksheets[0].getCell('B17').value;
    birthState = wb.worksheets[0].getCell('B18').value;
    birthCountry = wb.worksheets[0].getCell('B19').value;
    relationshipStatus = wb.worksheets[0].getCell('B20').value;
    // var = wb.worksheets[0].getCell('B21').value;
    // var = wb.worksheets[0].getCell('B22').value;
    // var = wb.worksheets[0].getCell('B23').value;
    // var = wb.worksheets[0].getCell('B24').value;
    // var = wb.worksheets[0].getCell('B25').value;
    // var = wb.worksheets[0].getCell('B26').value;
    // var = wb.worksheets[0].getCell('B27').value;
    // var = wb.worksheets[0].getCell('B28').value;
    immigrationOffice = wb.worksheets[0].getCell('B29').value;
    // var = wb.worksheets[0].getCell('B30').value;
    // var = wb.worksheets[0].getCell('B31').value;
    // var = wb.worksheets[0].getCell('B32').value;
    // var = wb.worksheets[0].getCell('B33').value;
    // var = wb.worksheets[0].getCell('B34').value;
    // var = wb.worksheets[0].getCell('B35').value;
    // var = wb.worksheets[0].getCell('B36').value;
    // var = wb.worksheets[0].getCell('B37').value;
    // var = wb.worksheets[0].getCell('B38').value;
    // var = wb.worksheets[0].getCell('B39').value;
    // var = wb.worksheets[0].getCell('B40').value;
    // var = wb.worksheets[0].getCell('B41').value;
    // var = wb.worksheets[0].getCell('B42').value;
})


async function main() {

    let driver = await new Builder().forBrowser("chrome").build();
    await driver.get("https://online.immi.gov.au/ola/app");
    console.log(username)


    await login();
    
    await newApplication();
    await page1();
    await page2();
    await page3();

    async function login() {
        await driver.findElement(By.name("username")).sendKeys(username);
        await driver.findElement(By.name("password")).sendKeys(password, Key.RETURN);
        await driver.findElement(By.name("continue")).sendKeys(Key.RETURN);

    }
    // async function nextPage() {
    //     driver.findElement(By.css("button[title='Go to next page']")).click();
    // }

    async function newApplication() {
        
        await driver.wait(until.titleIs("Online Account - My applications summary"), 1000);

        if ( driver.findElement(By.xpath("/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/button")).isDisplayed == false) {
            

            await driver.findElement(By.name("btn_newapp")).sendKeys(Key.RETURN);
            await driver.sleep(750)
            await driver.findElement(By.xpath("//div[2]/div/div[1]/div/div/div[12]/button")).click();
            await driver.sleep(500)
            await driver.findElement(By.xpath("//div/div[12]/div/button[3]")).click();
            await driver.sleep(3000)
        } else { 
            await driver.findElement(By.xpath("//div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/button")).click();

        }
 
    }

    async function continueApplication() {
        await driver.wait(until.titleIs("Online Account - My applications summary"), 1000);
        await driver.findElement(By.xpath("/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/button")).click();
        
    }
// Page 1

    async function page1() {

        if ( !driver.findElement(By.xpath("/html/body/form/div[1]/div/div/div[1]/section/div/div/div/div[5]/div/div/div/div/div/div/div[3]/div/div/div[2]/span/input")).checked ) {
        // if ( !driver.findElement(By.xpath("input[id*='input']")).checked ) {
            console.log("found it")
            await driver.findElement(By.css("input[id*='input']")).click();
            await driver.findElement(By.css("button[title='Go to next page']")).click();

        } else {
            await driver.findElement(By.css("button[title='Go to next page']")).click();

        }
    }

// Page 2
    async function page2(){
        if ( driver.findElement(By.xpath("/html/body/form/div[1]/div/div/div[1]/section/div/div/div/div[5]/div/div/div/div/div/div/div[15]/div/div[4]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).checked == false) {
            console.log("not checked page 2")

            await driver.findElement(By.xpath("//div[3]/div/div/div[2]/span/select")).sendKeys(currentLocation);
            await driver.findElement(By.xpath("//div[2]/div/div[4]/div/div/div[2]/span/select")).sendKeys(legalStatus);
            await driver.findElement(By.xpath("//div[5]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();
            await driver.findElement(By.xpath("//div[8]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();
            await driver.findElement(By.xpath("//div[11]/div/div[2]/div[2]/fieldset/div/label[1]/input")).click();
            await driver.findElement(By.xpath("//div[12]/div/button")).click();

            await driver.sleep(2000);
            await driver.findElement(By.xpath("//div[1]/div/div[2]/div/div/div[2]/div/div/div[1]/span/select")).sendKeys("letter for");

            await driver.findElement(By.xpath("//div[1]/div/div[4]/div/div/div[2]/span/input")).sendKeys("123");
            await driver.findElement(By.xpath("//div[1]/div/div[6]/div/div/div[2]/span/input")).sendKeys("123");
            await driver.findElement(By.xpath("//div[1]/div/div[9]/div/div[2]/div[2]/fieldset/div/label[2]/input")).click();
            await driver.findElement(By.xpath("//div[2]/div/div/div[2]/button")).click();

            await driver.sleep(1000);
            if (educationSector == "VET") {
                await driver.findElement(By.xpath("//div/div/div/div/div[14]/div/div[3]/div/div/div[2]/div/div/div[1]/span/select")).sendKeys("voca");
            } else {
                await driver.findElement(By.xpath("//div/div/div/div/div[14]/div/div[3]/div/div/div[2]/div/div/div[1]/span/select")).sendKeys("high");
            }
            
            await driver.findElement(By.xpath("//div/div/div/div/div[15]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();
            await driver.findElement(By.xpath("//div/div/div/div/div[15]/div/div[4]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();
            await driver.findElement(By.css("button[title='Go to next page']")).click();
            await driver.sleep(3000);
            await driver.findElement(By.xpath("/html/body/form/dialog/div/div/div/div[2]/div/div/div[2]/button")).click();

            } else {
                console.log("checked page 2")

                await driver.findElement(By.css("button[title='Go to next page']")).click();
                await driver.sleep(3000);
                await driver.findElement(By.xpath("/html/body/form/dialog/div/div/div/div[2]/div/div/div[2]/button")).click();
            }

    }
    
    
// Page 3

    // function ReadData(cell, row) {
    //     var excel = new ActiveXObject("npm.Application");
    //     var excel_file = excel.Workbooks.Open("C:\\RS_Data\\MyFile.xls");
    //     var excel_sheet = excel.Worksheets("Sheet1");
    //     var data = excel_sheet.Cells(cell, row).Value;
    //     document.getElementById('div1').innerText = data;
    // }
    async function page3(){
        if ( driver.findElement(By.xpath("//div/div/div/div[19]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).checked == false ) {

            await driver.sleep(3000);
            await driver.findElement(By.xpath("//div[1]/div/div[1]/div/div/div[2]/div/div/div[1]/span/input")).sendKeys(lastName);
            await driver.findElement(By.xpath("//div[1]/div/div[2]/div/div/div[2]/div/div/div[1]/span/input")).sendKeys(firstName);

            ///Gender
            if (gender == "Male") {
                await driver.findElement(By.xpath("//div[2]/div/div/div[2]/fieldset/div/label[2]/input")).click();
            } else {
                await driver.findElement(By.xpath("//div[2]/div/div/div[2]/fieldset/div/label[1]/input")).click();
            }

            ///Date of Birth
            await driver.findElement(By.xpath("//div[3]/div/div/div[2]/div/input")).sendKeys(dateOfBirth, Key.TAB);
            await driver.sleep(100);

            
            ///Passport Number
            await driver.findElement(By.xpath("//div/div/div/div/div/div/div[7]/div/div[1]/div/div/div[2]/span/input")).sendKeys(passportNumber);
            

            ///Country of Passport
            await driver.findElement(By.xpath("//div[2]/div/div/div[2]/span/select")).sendKeys(passportCountry);
            

            ///Nationality of Passport
            await driver.findElement(By.xpath("//div[3]/div/div/div[2]/span/select")).sendKeys(passportNationality);

            ///Date of Passport Issuance
            await driver.findElement(By.xpath("//div[4]/div/div/div[2]/div/input")).sendKeys(passportIssuanceDate, Key.TAB);

            ///Date of Passport Expiry
            await driver.findElement(By.xpath("//div/div/div/div[7]/div/div[5]/div/div/div[2]/div/input")).sendKeys(passportExpiryDate, Key.TAB);
            
            ///Issuer of Passport, Place of Issue, Issuing Authority
            await driver.findElement(By.xpath("//div/div/div/div/div[7]/div/div[7]/div/div/div[2]/span/input")).sendKeys(passportIssuer);
            /// Does this applicant have a national identity card?
            await driver.findElement(By.xpath("//div[9]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();

            /// Place of Birth
            await driver.findElement(By.xpath("//div/div/div/div[10]/div/div[2]/div/div/div[2]/span/input")).sendKeys(birthCity);
            await driver.findElement(By.xpath("//div/div/div/div[10]/div/div[3]/div/div/div[2]/span/input")).sendKeys(birthState);
            await driver.findElement(By.xpath("//div/div/div/div[10]/div/div[4]/div/div/div[2]/div/div/div[1]/span/select")).sendKeys(birthCountry);

            /// Relationship Status
            await driver.findElement(By.xpath("//div/div[12]/div/div[2]/div/div/div[2]/div/div/div[1]/span/select")).sendKeys(relationshipStatus);

            /// Other Names
            // if No then click
            await driver.findElement(By.xpath("//div/div/div[13]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();
            await driver.sleep(1500);


            /// Citizenship in Passport
            await driver.findElement(By.xpath("//div[4]/div/div/div/div/div/div/div[14]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[1]/input")).click();

            
            /// Citizen in Other Country?
            await driver.findElement(By.xpath("//div/div/div[14]/div/div[3]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();

            /// Other Passports?
            await driver.findElement(By.xpath("//div/div/div[15]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();


            /// Other IDs?
            await driver.findElement(By.xpath("//div/div[16]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();

            /// Health Exam?
            await driver.findElement(By.xpath("//div/div/div/div[19]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();

            await driver.findElement(By.css("button[title='Go to next page']")).click();
        } else {
            await driver.findElement(By.css("button[title='Go to next page']")).click();
        }
    }
    

// /// Page 4    
//     await driver.sleep(1500);

//     await driver.findElement(By.xpath("//div/div[10]/div/div/div[2]/div/div/div[1]/fieldset/div/label[1]/input")).click();
//     await driver.findElement(By.css("button[title='Go to next page']")).click();

// /// Page 5
//     /// Accompanying Members of the Family Unit
//     await driver.sleep(1500);
//     await driver.findElement(By.xpath("//div[5]/div/div/div/div/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/fieldset/div/label[2]/input")).click();
//     await driver.findElement(By.css("button[title='Go to next page']")).click();

// /// Page 8
//     /// Usual Country of Residence
//     await driver.findElement(By.xpath("//div[5]/div/div/div/div/div/div/div[2]/div/div[2]/div/div/div[2]/div/div/div[1]/span/select")).sendKeys(passportCountry);
    
//     await driver.findElement(By.xpath("//div[5]/div/div/div/div/div/div/div[3]/div/div[4]/div/div/div[2]/div/div/span/input")).sendKeys(immigrationOffice);
}


main();




