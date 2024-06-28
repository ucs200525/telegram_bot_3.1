const puppeteer = require('puppeteer');

(async () => {
    const [,, city, date, outputImagePath] = process.argv;

    const browser = await puppeteer.launch(); // Launch browser
    const page = await browser.newPage();
    await page.goto('https://www.drikpanchang.com/muhurat/panchaka-rahita-muhurat.html');

    // Clear the existing city and date values
    await page.evaluate(() => {
        document.getElementById('dp-direct-city-search').value = '';
        document.getElementById('dp-date-picker').value = '';
    });

    // Input the new city and date values
    await page.type('#dp-direct-city-search', city);
    await page.type('#dp-date-picker', date);

    // Screenshot the specified element
    const element = await page.$('.dpMuhurtaCard.dpFlexEqual');
    await element.screenshot({ path: outputImagePath });

    await browser.close(); // Close browser
})();
