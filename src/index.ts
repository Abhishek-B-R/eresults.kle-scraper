import puppeteer from "puppeteer";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs"

interface dataProps{
    yearOfJoining:number,
    branchInTwoChars:string,
    startUSN:number,
    endUSN:number
}

const dirname=__dirname.split("/").filter((e)=>e!=="dist").join("/")
const filePath = path.join(dirname, 'data.json');
const data:dataProps = JSON.parse(fs.readFileSync(filePath, 'utf8'));

function delay(time:number) {
    return new Promise(resolve => setTimeout(resolve, time));
}

// Create a new Excel workbook and worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Results");

worksheet.columns = [
    { header: "Name", key: "name", width: 30 },
    { header: "USN", key: "usn", width: 20 },
    { header: "CGPA", key: "cgpa", width: 10 },
];

(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox','--start-maximized']
      });
      
    const [page] = await browser.pages();

    const { width, height } = await page.evaluate(() => {
    return {
        width: window.screen.width,
        height: window.screen.height
    };
    });

    await page.setViewport({ width, height });

    await page.goto("https://eresults.kletech.ac.in/", { waitUntil: 'networkidle2' });

    const noOfUSN=data.endUSN;

    for(let i =data.startUSN;i<=noOfUSN;i++){
        try{
            await page.evaluate(() => {
                const input = document.querySelector('#usn');
                if (input) (input as HTMLInputElement).value = ''; 
            });
    
            await page.click('#usn');
            await page.keyboard.down('Control');
            await page.keyboard.press('A');
            await page.keyboard.up('Control');
            await page.keyboard.press('Backspace'); 

            const usnPrototype=`01FE${data.yearOfJoining}B${data.branchInTwoChars}`
    
            if(i<10)        await page.type('#usn', `${usnPrototype}00`+i);
            else if(i<100)  await page.type('#usn', `${usnPrototype}0`+i);
            else            await page.type('#usn', `${usnPrototype}`+i);
    
            if(i<=data.startUSN)    await delay(10000)
            page.click('button[type="submit"][class="myButton"][formaction*="index.php?option=com_examresult&task=getResult"]')
            await delay(3000)
            
            const name = await page.evaluate(() => {
                const element = document.querySelector('.uk-card.stu-data.stu-data1');
                return element ? element.textContent?.trim() : null;
            });
            
            let usnVal=await page.evaluate(()=>{
                const element = document.querySelector('.uk-card.stu-data.stu-data2')
                return element?.textContent?.trim();
            })
            usnVal=usnVal?.split('\n')[0].trim();
            
            let CGPA=await page.evaluate(()=>{
                const element = document.querySelectorAll('.uk-card.uk-card-default.uk-card-body.credits-sec1')
                return element[3]?.textContent?.trim();
            })
            CGPA=CGPA?.split('\n')[1].trim();
    
            
            if(name && usnVal && CGPA){
                console.log(name,usnVal,CGPA)
                worksheet.addRow({ name, usn: usnVal, cgpa: CGPA });
            }
            await page.goBack();
        }catch(e){
            console.error(`Error on page ${i}:`, e);
            await page.goBack();
        }
    }
    await workbook.xlsx.writeFile("results.xlsx");
    console.log("Excel file saved as results.xlsx");

    await browser.close();
})()