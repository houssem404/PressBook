# PressBook
This is a JavaScript code that uses the fs module to read a list of links from a text file, puppeteer module to automate a headless browser (Google Chrome) to capture screenshots of the web pages, and the officegen module to generate a PowerPoint presentation.

The code uses the puppeteer.launch() function to open a new browser instance, creates a new page using browser.newPage(), navigates to the desired web page using page.goto(), and captures a screenshot of the entire page using page.screenshot(). It then closes the browser using browser.close(). The captured screenshot is then added to a new slide in the PowerPoint presentation using the slide.addImage() function.

The code loops through each link in the list, captures a screenshot of the corresponding page, adds the screenshot to a new slide in the presentation, and adds the site name, date, and compressed link as text to the same slide.

Finally, the PowerPoint presentation is saved to a file using the pptx.generate() function and passing a write stream created by fs.createWriteStream().

Install  fs , puppeteer,officegen  
