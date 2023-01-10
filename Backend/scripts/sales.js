const puppeteer = require("puppeteer");
(async () => {
    const browser = await puppeteer.launch({headless: false, defaultViewport: null, args: [`--start-maximized`]});
    const page = await browser.newPage();
    await page.goto('https://decroixintl-my.sharepoint.com/personal/adill_decroixintl_com/_layouts/15/onedrive.aspx?FolderCTID=0x01200044A0C9535FFEE34A9857240097601E25&id=%2Fpersonal%2Fadill%5Fdecroixintl%5Fcom%2FDocuments%2FSales%20and%20QC%2FDaily%20Logs%2F01%20January%202023');
    await page.screenshot({path: "testing.png"});
    await page.waitForSelector('input[name=loginfmt]');
    await page.$eval('input[name=loginfmt]', el => el.value = 'mbreden@decroixintl.com');
    await page.waitForSelector("#idSIButton9")
    await page.click("#idSIButton9");
    await page.waitForTimeout(3000);
    await page.waitForSelector('input[name=passwd]');
    await page.$eval('input[name=passwd]', el => el.value = 'Wabbawinner123');
    await page.waitForTimeout(3000);
    await page.click(
        "#idSIButton9"
    );
    //await page.waitForTimeout(3000);
    await page.waitForSelector("#idBtn_Back")
    await page.click("#idBtn_Back");
    //await page.waitForTimeout(3000);
    await page.waitForSelector('div[title="January 2023 Sales & QC Log.xlsx"]');
    await page.click('div[title="January 2023 Sales & QC Log.xlsx"]');
    await page.waitForSelector('Button[name="Download"]')
    await page.click('Button[name="Download"]');
    await page.waitForTimeout(2000);
    // const dimensions = await page.evaluate(() => {
    //     return {
    //       width: document.documentElement.clientWidth,
    //       height: document.documentElement.clientHeight,
    //       deviceScaleFactor: window.devicePixelRatio,
    //     };
    // });

    // console.log('Dimensions:', dimensions);
    await browser.close();
})();