; Setup: Create a new instance of Chrome
CreateChromeInstance()
{
    chrome := ComObject("Selenium.ChromeDriver")
    return chrome
}

; Example 1: Navigate to a URL
NavigateToUrl(chrome, url)
{
    chrome.Get(url)
}

; Example 2: Find and interact with elements using different locators
InteractWithElements(chrome)
{
    ; Click a button by ID
    chrome.findElementById("submit-button").Click()

    ; Type into an input field by name
    chrome.findElementByName("search").SendKeys("AutoHotkey Selenium")

    ; Select an option from a dropdown by CSS selector
    dropdown := chrome.findElementByCssSelector("select#country")
    dropdown.findElementByCssSelector("option[value='US']").Click()
}

; Example 3: Wait for elements to be present
WaitForElement(chrome, locator, timeout := 10000)
{
    try {
        chrome.Manage().Timeouts().ImplicitWait := timeout
        element := chrome.findElement(locator)
        return element
    } catch {
        MsgBox, Element not found within the specified timeout.
        return false
    }
}

; Example 4: Handle alerts
HandleAlert(chrome, action := "accept")
{
    try {
        alert := chrome.SwitchTo().Alert()
        if (action = "accept")
            alert.Accept()
        else if (action = "dismiss")
            alert.Dismiss()
        else if (action = "getText")
            return alert.Text()
    } catch {
        MsgBox, No alert present.
    }
}

; Example 5: Execute JavaScript
ExecuteJavaScript(chrome, script)
{
    return chrome.ExecuteScript(script)
}

; Example 6: Take a screenshot
TakeScreenshot(chrome, filename)
{
    screenshot := chrome.GetScreenshot()
    screenshot.SaveAsFile(filename, 2) ; 2 is for PNG format
}

; Example 7: Manage cookies
ManageCookies(chrome, action := "getAll")
{
    if (action = "getAll")
        return chrome.Manage().Cookies().GetCookies()
    else if (action = "deleteAll")
        chrome.Manage().Cookies().DeleteAllCookies()
    else if (action = "add") {
        cookie := {"name": "exampleCookie", "value": "exampleValue"}
        chrome.Manage().Cookies().AddCookie(cookie)
    }
}

; Usage example
chrome := CreateChromeInstance()
NavigateToUrl(chrome, "https://www.example.com")
InteractWithElements(chrome)
WaitForElement(chrome, {"using": "id", "value": "results"})
HandleAlert(chrome, "accept")
ExecuteJavaScript(chrome, "return document.title;")
TakeScreenshot(chrome, "screenshot.png")
ManageCookies(chrome, "getAll")
