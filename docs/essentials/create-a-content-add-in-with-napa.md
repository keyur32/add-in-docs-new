
# Create a content add-in for Excel with Napa Office 365 Development Tools


You can also build task pane and content add-ins using [Visual Studio](http://msdn.microsoft.com/en-us/library/a23c5ce8-6de3-40f6-a86a-85d3592bef3e%28Office.15%29.aspx) or a [text editor](http://msdn.microsoft.com/en-us/library/d5411d35-9ef6-4e21-ba2b-4d2b1ee81359%28Office.15%29.aspx). If you're not sure which tool to use, see [Development basics](../../docs/overview/platform-overview.md#StartBuildingApps_DevelopmentBasics).

For more information about Napa, see [Create Office Add-ins with Napa Office 365 Development Tools](../../docs/essentials/create-office-add-ins-with-napa.md).


## Prerequisites



- A [Microsoft account](http://www.microsoft.com/en-us/account/default.aspx)
    
- The URL for the [Napa Office 365 Development Tools](https://www.napacloudapp.com/ ) web app
    

## Create a basic Office Add-in



1. Open the [Napa Office 365 Development Tools](https://www.napacloudapp.com/ ) web app in your browser and sign in using your Microsoft account credentials.
    
2. Choose the  **Add New Project** tile.
    
    The  **Add New Project** tile appears only if you've created other projects. If this is your first project, skip to the next step.
    

    **New project tile**

    ![Projects page](../../images/08fc36cf-7cc1-442f-a9a5-b6bb30d786a4.png)

3. Choose the  **Content Add-in for Office** tile and name the projectMyFirstContentAddin. Choose the default  **Basic add-in** option and then choose the **Create** button.
    
    **Content add-in tile**

    ![Excel app tile](../../images/Apps_NAPA_Excel_Tile.png)
    <br>
    The code editor opens and shows the default webpage, which already contains some sample code that you can run without doing anything else.
    

### Run the sample Office Add-in


1. On the side of the page, choose the  **Run** button
![Run button](../../images/Apps_NAPA_Run_Button.png).
    
    Excel Online opens, and the sample Office Add-in appears. You can experiment with its features by choosing  **Edit Workbook > Edit in Excel Online**.
    
2. When you are ready to move on, close Excel Online.
    

## Add HTML and JavaScript files to the project



1. In Napa, choose the  **New folder** button.
    
2. Name the folder  **MyAddinPage**.
    
3. Open the shortcut menu for the  **MyAddinPage** folder (right-click the folder), and then choose **Add new file**.
    
    The  **New File** dialog box opens.
    
4. Choose the  **Html Page** tile, name the fileMyAddinPage, and then choose the  **Create** button.
    
5. Open the shortcut menu for the  **MyAddinPage** folder and then choose **Add new file**.
    
    The  **New File** dialog box opens.
    
6. Choose the  **JavaScript File** tile, name the fileMyAddinPage, and then choose the  **Create** button.
    
    Next, we'll modify the look and feel of the add-in and point it to the HTML page you just created.
    

## Change the add-in properties



1. On the side of the page, choose the  **Properties** button
![Properties button](../../images/Apps_Napa_Properties_Button.png).
    
    The properties of the Office Add-in appear.
    
2. Set the following properties:
    
      -  **Name** property asMy First Content Add-in
    
  -  **StartPage** asMyAddinPage/MyAddinPage.html
    
  -  **Description** asThis add-in gets data from a cell and writes data to a cell. This add-in also responds to events in the worksheet.
    
  -  **Initial width** as 520
    
  -  **Initial height** as 400
    

    The  **Name** and **Description** properties help users understand the purpose of the add-in when it appears in a list of available add-ins for an Office application. The size properties specify how much space the add-in requires. The **Start Page** property points to the page that appears in the add-in when you start the project.
    
3. Choose the  **Apply** button at the bottom of the **Properties** page, and then choose the **Explore** button
![Explore button](../../images/Apps_Napa_Explore_Button.png) on the left toolbar. This saves the property values and opens the Explore page.
    
     >**Note**  The  **Properties** editor shows the most common settings of an Office Add-in. It doesn't show all of the possible settings of an Office Add-in. If your scenario requires you to modify settings that don't appear in the **Properties** editor, you can create your add-in by using [Visual Studio](http://msdn.microsoft.com/en-us/library/a23c5ce8-6de3-40f6-a86a-85d3592bef3e%28Office.15%29.aspx) or a [text editor](http://msdn.microsoft.com/en-us/library/7aac2fdc-1a04-45ec-a1dc-da26e646a364%28Office.15%29.aspx). 

## Get data from a worksheet


Your Office Add-in can get the value of a single cell or the values of a collection of cells. The most basic task here is to get the value of a single cell that a user chooses in a worksheet. After completing these steps, you choose a cell in Excel and then choose a button in the add-in - the data from the cell that you chose appears in a control in the add-in.


1. On the side of the page, choose  **MyAddinPage.html**.
    
    The MyAddinPage webpage appears in the code editor.
    
2. Replace all the code within the  `<head>` tags (including the opening and closing `<head>` tags) with this code.
    
```HTML
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title></title>
    <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>
    
    <link href="../Content/Office.css" rel="stylesheet" type="text/css" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    
    <link href="../App/App.css " rel="stylesheet" type="text/css" />
    <script src="../App/App.js" type="text/javascript"></script>
    
    <script src="MyAddinPage.js" type="text/javascript"></script>
</head>

```


    This code gives your MyAddinPage webpage the same JavaScript libraries and CSS file references as the default Home.html file. The following table describes each file reference.
    

|**File**|**Description**|
|:-----|:-----|
|**App.css, Office.css**|The default CSS files of the project. You can use these pages to define certain aspects of how the webpage appears.|
|**MyAddinPage.js**|The JavaScript file that you created for your page. |
|**App.js**|Located in the  **Add-in** folder of the project, App.js is the default JavaScript file of the add-in, and contains some example code to get you started.|
3. Replace the opening and closing  `<body>` tags with this code.
    
    This code adds all of the controls that we'll use in this walkthrough. It also adds a table that will contain stock data related to symbols that you add to the worksheet. 
    


```HTML
  <body>
<div style="padding: 15px; overflow: auto; border: .2em solid #000;">

<table>
<tr>
<td>

<button id="get-text" style="width: 100px;">Get symbol</button>
</td>
<td>
<button id="bind-text" style="width: 100px;">Bind to cell</button>
</td>
</tr>
<tr>
<td>
<input id="input" style="width: 100px;"/>
</td>
<td>
<button id="add-text" style="width: 100px;">Add symbol</button>
</td>
</tr>

</table>
<h1><div id="stock-name"></div></h1>
<table border="true">
<tr>
<td>
<table>
<tr>
<td>Prev close:</td>
<td id="prev-close"></td>
</tr>

<tr>
<td>Open:</td>
<td id="open"></td>
</tr>

<tr>
<td>Bid:</td>
<td id="bid"></td>
</tr>
<tr>
<td>Ask:</td>
<td id="ask"></td>
</tr>
<tr>
<td>1y Target Est:</td>
<td id="target-est"></td>
</tr>
<tr>
<td>Days range:</td>
<td id="days-range"></td>
</tr>
</table>
</td>
<td>
    <table>
<tr>
<td>Volume:</td>
<td id="volume"></td>
</tr>

<tr>
<td>Avg daily volume:</td>
<td id="avg-volume"></td>
</tr>

<tr>
<td>Market capitalization:</td>
<td id="market-cap"></td>
</tr>
<tr>
<td>PE Ratio:</td>
<td id="pe-ratio"></td>
</tr>
<tr>
<td>Earnings p share:</td>
<td id="earnings"></td>
</tr>
<tr>
<td>Dividend yield:</td>
<td id="yield"></td>
</tr>
</table>
</td>
</tr>
</table>
</div>

</body>



```

4. Open the MyAddinPage.js file, and then add this code.
    
    When you run the code, you'll add a stock symbol to a cell. The code gets that stock symbol and shows data related to that symbol in a table.
    


```js
  /// <reference path../../Scripts/App.js" />

(function () {
    "use strict";
    
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {


$('#get-text').click(getTextFromDocument);      
        });
    }

})();
function getTextFromDocument() {

    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        { valueFormat: "unformatted", filterType: "all" },

        function (asyncResult) {
            showStockData(asyncResult.value);
        });

}
function showStockData(symbol){
    // Yahoo YQL - http://developer.yahoo.com/yql/ 
var yql = 'select * from yahoo.finance.quotes where symbol in (\'' + symbol + '\')';
var queryURL = 'https://query.yahooapis.com/v1/public/yql?q=' + yql + '&amp;format=json&amp;env=http%3A%2F%2Fdatatables.org%2Falltables.env&amp;callback=?';

$.getJSON(queryURL, function(results) {
if(results.query.count > 0)
{
var quotes = results.query.results.quote;

$('#stock-name').text(quotes.Name);
$('#prev-close').text(quotes.PreviousClose);
$('#open').text(quotes.Open);
$('#bid').text(quotes.Bid);
$('#ask').text(quotes.Ask);
$('#target-est').text(quotes.OneyrTargetPrice);
$('#days-range').text(quotes.DaysRange);
$('#volume').text(quotes.Volume);
$('#avg-volume').text(quotes.AverageDailyVolume);
$('#market-cap').text(quotes.MarketCapitalization);
$('#pe-ratio').text(quotes.PERatio);
$('#earnings').text(quotes.EarningsShare);
$('#yield').text(quotes.DividendYield);

}

});

}



```


    The call to  `getSelectedDataAsync` passes an anonymous function with a parameter, named `asyncResult`, as the callback argument. When the callback function runs, it uses the  `asyncResult` parameter to access the value property of the `AsyncResult` object to display the data in the cell that the user chose.
    
     >**Note**  Like other methods in the JavaScript API for Office, this method is asynchronous in that it doesn't prevent the add-in from performing other operations while this method runs.

### Run it!


1. On the side of the page, choose the  **Run** button
![Run button](../../images/Apps_NAPA_Run_Button.png).
    
    Excel Online opens, and the Office Add-in appears. Choose  **Edit Workbook > Edit in Excel Online**.
    
2. Enter MSFT into any cell.
    
    This abbreviation is the stock-ticker symbol for Microsoft.
    
3. In the Office Add-in, choose the  **Get symbol** button.
    
    Data related to the ticker symbol MSFT appear in a table.
    

    **Data appearing in a table of the add-in**

    ![MSFT appears in the app when you press the button](../../images/Apps_NAPA_Excel_Get.png)
    <br>
    This example shows how to get data from a cell. In your add-in, you might use that technique to look up information in a database, get information from another service, or perform a calculation. You could add code to perform those sorts of operations to the anonymous function that you pass as a parameter to the  `getSelectedDataAsync` method.
    
4. Close Excel Online.
    
    In the next example, you'll take data that the user enters into a control on the add-in, and you'll put that data into a cell in the worksheet.
    

## Put data into selected cells in a worksheet


Your Office Add-in can put data into any cell or collection of cells. The most basic task here is to put the data into a cell that a user chooses in a worksheet. After you complete these steps, a user can add text to a cell in the worksheet by choosing a button in the add-in.


1. In the code editor, open the  **MyAddinPage.js** file, and then add this code.
    
```js
  function addTextToDocument() {

    var e = document.getElementById("input");
    var text = e.value;

    Office.context.document.setSelectedDataAsync(text,
        function (asyncResult) {});
}

```


    This code gets the text from the  `input` text box in MyAddinPage.html and places that text into a cell that the user chooses in the worksheet.
    
2. Replace the  `initialize` function with this code.
    
```js
  Office.initialize = function (reason) {
        $(document).ready(function () {
$('#get-text').click(getTextFromDocument);
$('#add-text').click(addTextToDocument);
            
        });
    }

```


### Run it!


1. On the side of the page, choose the  **Run** button
![Run button](../../images/Apps_NAPA_Run_Button.png).
    
    Excel Online opens, and the Office Add-in appears.
    
2. In Excel Online, choose any cell.
    
3. In the add-in, enter  **YHOO** in the text box next to the **Add Symbol** button, and then choose the **Add symbol** button.
    
    The text  **YHOO** appears in the cell that you chose.
    

    **The text MSFT appearing in the selected cell**

    ![MSFT appears in a cell when you press the button](../../images/Apps_Napa_Put.png)
    <br>
    This example is simple, but it shows how to put data into a cell. Your Office Add-in might use a stock service to get the closing price of a stock, and then add that price to a cell, which might perform other calculations.
    
4. Close Excel Online. 
    

## Handle an event in a worksheet


So far, your Office Add-in requires the user to choose a button to get and set data. By doing a couple of more steps, you can also get and set data automatically when a user chooses a cell.


1. In the code editor, open the  **MyAddinPage.js** file, and then replace the `initialize` function with this code.
    
```js
  Office.initialize = function (reason) {
        $(document).ready(function () {
$('#get-text').click(getTextFromDocument);
$('#add-text').click(addTextToDocument);
        Office.context.document.addHandlerAsync
        (Office.EventType.DocumentSelectionChanged, updateApp);
            
        });
    }

```


    This code binds functions to the buttons on the page, and adds an event handler that's called when the user chooses a cell.
    
2. Add this code to the MyAddinPage.js file.
    
```js
  function updateApp()
{
        getTextFromDocument();
}

```


    This method is called when a user chooses a cell. The code calls the method that you defined earlier. That method gets the value of the chosen cell (stock symbol) and shows the data related to that symbol in a table.
    

### Run it!


1. On the side of the page, choose the  **Run** button
![Run button](../../images/Apps_NAPA_Run_Button.png). 
    
    Excel Online opens, and the Office Add-in appears. Choose  **Edit Workbook > Edit in Excel Online**.
    
2. In the add-in, enter  **MSFT** in the text box next to the **Add symbol** button, and then choose the **Add symbol** button.
    
3. Choose another cell, and then choose the cell that contains  **MSFT**.
    
    Ticker data for the symbol  **MSFT** appears in the table.
    
4. Close Excel Online.
    

## Bind to cells in a worksheet


The most advanced way to get and set data is to establish a binding with a cell or a collection of cells in a worksheet. You can prompt users to choose the cells that they want the add-in to use. Then, you can get data from those cells or put data into those cells at any time.


1. In the code editor, open the MyAddinPage.js file, and then add this code. This code establishes a binding to a cell that the user chooses. This code also defines a method to call when the data in the bound cell changes.
    
```js
  function addBindingFromSelection() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' },
        function (asyncResult) {
            Office.select("bindings#MyBinding").addHandlerAsync
               (Office.EventType.BindingDataChanged, onBindingSelectionChanged);
        }
    );
}


function onBindingSelectionChanged(eventArgs) {

    Office.select("bindings#MyBinding").getDataAsync
        (function (asyncResult) {

            if (asyncResult.value !== "") {
                showStockData(asyncResult.value);
            }

         });
}

```

2. Replace the  `initialize` function with this code.
    
```js
  Office.initialize = function (reason) {
        $(document).ready(function () {
$('#get-text').click(getTextFromDocument);
$('#add-text').click(addTextToDocument);
$('#bind-text').click(addBindingFromSelection);

        Office.context.document.addHandlerAsync
        (Office.EventType.DocumentSelectionChanged, updateApp);

        });
    }

```


### Run It!


1. On the side of the page, choose the  **Run** button
![Run button](../../images/Apps_NAPA_Run_Button.png).
    
    Excel Online opens, and the Office Add-in appears. Choose  **Edit Workbook > Edit in Excel Online**.
    
2. In Excel Online, select any cell. Then, in the Office Add-in, choose the  **Bind to cell** button.
    
3. In the add-in, enter  **MSFT** in the text box next to the **Add Symbol** button, and then choose the **Add symbol** button.
    
    The text  **MSFT** appears in the cell that you chose. Because the value of the cell changed, data related to that cell appears in the table.
    

    **Table shows data for the ticker symbol MSFT**

    ![Shows binding to a cell](../../images/Apps_Napa_Bind.png)

4. Close Excel Online. 
    

## Debug your content add-in in Internet Explorer


If you start your add-in in Excel Online, and you use Internet Explorer (IE), you can use F12 developer tools to debug the JavaScript, HTML, and Cascading Style Sheets (CSS) of your content add-in. 

Here's how to open F12 tools, start the debugger, and force execution to stop on a line of code in your MyAddinPage.js file.


1. On the side of the page, choose the  **Run** button
![Run button](../../images/Apps_NAPA_Run_Button.png).
    
    Excel Online opens, and the Office Add-in appears. Choose  **Edit Workbook > Edit in Excel Online**.
    
2. Press the F12 key on your keyboard.
    
    The F12 tools open in a separate window.
    
3. In the F12 tools window, open the  **Debugger** tab.
    
4. Use the Ctrl-O keyboard shortcut to open a document, and then enter MyAddinPage.js in the filter text box.
    
    The contents of the MyAddinPage.js file appears in the window.
    
5. Set a breakpoint on the  `addTextToDocument` method.
    
    For more information about how to set a breakpoint in the F12 tool window, see [Breaking Code Execution](http://go.microsoft.com/fwlink/?LinkID=267272).
    
6. In the add-in, enter  **MSFT** in the text box next to the **Add Symbol** button, and then choose the **Add symbol** button.
    
    In the F12 tools window, execution stops on the  `addTextToDocument` method.
    
    See [Using the F12 developer tools](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29) for more information.
    
    If you use a browser other than Internet Explorer, search your browser documentation.
    

## Next Steps


Now that you've created a basic content add-in for Excel, you can:


- Share your project with someone by choosing the  **Share Project** button
![The Share Project button](../../images/NAPA_Apps_Share.png). Napa creates a copy of your project and gives you a public link that you can give to anyone.
    
- Publish your add-in by choosing the  **Publish** button
![Publish button](../../images/Apps_NAPA_Publish.png).
    
    For more information, see [Package your add-in using Napa or Visual Studio to prepare for publishing](../publish/package-your-add-in-using-napa-or-visual-studio.md).
    
- Open your project in Visual Studio by choosing the  **Open in Visual Studio** button
![Open in Visual Studio button](../../images/Apps_Napa_OpenInVS.png). Napa automatically installs the necessary tools and opens your project in Visual Studio.
    
- Create a task pane add-in for Excel by using Visual Studio. For more information, see [Create a task pane or content add-in with Visual Studio](../../docs/essentials/create-a-task-pane-or-content-add-in-with-visual-studio.md).
    
- Learn more about Office Add-ins in the [Office Add-ins platform overview](../../docs/overview/platform-overview.md).
    

## Additional resources



- [Understanding the JavaScript API for Office](../../docs/overview/understanding-the-javascript-api-for-office.md)
    
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
    
