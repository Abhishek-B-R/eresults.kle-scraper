import puppeteer from "puppeteer";
import ExcelJS from "exceljs";

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

    const noOfUSN=120;

    for(let i =1;i<=noOfUSN;i++){
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
    
            if(i<10)        await page.type('#usn', "01FE23BCI00"+i);
            else if(i<100)  await page.type('#usn', "01FE23BCI0"+i);
            else            await page.type('#usn', "01FE23BCI"+i);
    
            if(i<=1)    await delay(10000)
            page.click('button[type="submit"][class="myButton"][formaction*="index.php?option=com_examresult&task=getResult"]')
            await delay(2000)
            
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