# Overview 
XL2SO (Excel to ScriptableObject) is a Unity Editor Extension that helps to create ScriptableObject Script (.cs) and ScriptableObject Instance (.asset) from Excel spreadsheet.  
The purpose of this extension is to minimize the amount of time spent on writing script code, generating instance and setting those parameters.  
In comparison to other similar extensions,  XL2SO has the following <span style="color: blue; ">advantages</span>.

<span style="color: blue; font-weight: bold;"> Convenience: </span>
- No need to create extra setting files and templates. 
- No need to change existing Excel file.

<span style="color: blue; font-weight: bold;"> Simplicity: </span>
- UI consists of bare minimum components and it's easy to understand how to use.

<span style="color: blue; font-weight: bold;"> Flexibility: </span>
- Provides an interface to select interested cells, specify data type and accessibility.
- Independent generation feature of both script and instance. You can update a generated script
(e.g., add fields/methods and inherit other ScriptableObject script) and use it for instance generation feature.  

# Features 
XL2SO provides 2 features: 1) Generate ScriptableObject Script 2) Generate ScriptableObject Instance.
Both features need just 3 steps shown below:

## Generate ScriptableObject Script
![](/Images/Features/SG.svg)
<u><b>Steps:</b></u>
1. Configure Excel file settings.
2. Specify the following properties for each field.
    - Enable/Disable
    - Accessibility
    - Data Type
3. Configure output settings. 

## Generate ScriptableObject Instance
![](/Images/Features/IG.svg)
<u><b>Steps:</b></u>
1. Specify ScriptableObject Script. 
2. Configure Excel file settings.
3. Configure Output settings.

# Getting Started 
This chapter describes example usage with a sample file named <b> SampleBook.xlsx</b> which contains <b>Item</b> sheet below.
![](/Images/GettingStarted/ExcelSheet.svg)

## Installation
- Import XL2SO.unitypackage
or
- Add files and directories in "Assets/" directory to your project.

## Launch XL2SO
Click <b>[Tools]</b> > <b>[XL2SO]</b> in Unity menu bar.
![](/Images/GettingStarted/MenuItem.png)

Then the following window is opened.
![](/Images/GettingStarted/Menu.png)

## Generate ScriptableObject Script
Click <b>[Generate ScriptableObject Script]</b> icon in the startup window.

![](/Images/GettingStarted/SG.svg) 

<img src = "/Images/GettingStarted/1.svg" style ="vertical-align:top;"> Excel Settings
- Set Excel file to <b> Book </b> field.
    - Then XL2SO automatically set <b> Sheet </b> and <b> Cells </b> fields.
- If there are some sheets in the Excel file, select a sheet which contains target data.
- If you need to change a range of interested cells, click "Select Cells" button and adjust the range in Excel Viewer.
    - Please correspond first row in the range to column label because XL2SO extracts it for field name.
    - In other words, it needs to select only single line of column labels. Including other lines which contain each data is optional for predicting its data type shown in Field List.

![](/Images/GettingStarted/EVWindow.svg)

<img src = "/Images/GettingStarted/2.svg" style ="vertical-align:top;"> Field List Settings 
- Specify Ignore or not, Accessibility and Data Type for each field.
    - Data Type popup shows primary data type of C# and user defined enums.
    - If other data type like class, unity object and undefined enum is needed, please manually update the code after script generation.
    - In this example case, EffectTypes enum is prepared in advance so it's visible in the above screenshot.
    ![](/Images/GettingStarted/EffectTypes.png)

<img src = "/Images/GettingStarted/3.svg" style ="vertical-align:top;"> Output Settings
- Specify output directory and file name.
- Click "Create" button and then XL2SO generates a script. 

![](/Images/GettingStarted/SG_Result.png)

<img src = "/Images/GettingStarted/4.svg" style ="vertical-align:top;"> Return to Menu Button
- Click "Return to Menu" button to return to startup menu.

## Generate ScriptableObject Instance

Click <b>[Generate ScriptableObject Instance]</b> icon in the startup window.

![](/Images/GettingStarted/IG.svg) 

<img src = "/Images/GettingStarted/1.svg" style ="vertical-align:top;"> ScriptableObject Script Settings
- Specify ScriptableObject Script to <b> Script </b> field.

<img src = "/Images/GettingStarted/2.svg" style ="vertical-align:top;"> Excel Settings
- Configure Excel settings as same as <b>Generate ScriptableObject Script</b> feature.

<img src = "/Images/GettingStarted/3.svg" style ="vertical-align:top;"> Instance List Settings
- Select instance naming rule from the following 2 options.
    - <b> Base Name With Index </b>: Name each instance like [BaseName], [BaseName] (1), [BaseName] (2),...
    - <b> Field Value</b>: Use each value of a specified field.

<img src = "/Images/GettingStarted/4.svg" style ="vertical-align:top;"> Output Settings
- Specify output directory.
- Click "Generate" button and then XL2SO generates instances.

![](/Images/GettingStarted/IG_Result.png) 

# Requirement
- .Net Framework 4.5 or above

# Support Excel Format
- .xlsx (Excel 2007 - )
- .xls  (Excel 97 - 2003)

# Test Environment
- Unity 2019.3.6f1
- .Net Standard 2.0

# Licence
- MIT

<style>
  h1 {
    border-bottom-width: 1px;
    border-bottom-style: solid;
    border-bottom-color: black;
    font-size: 30px;
  }
  h2 {
    padding: 6px;
    line-height: 36px;
    font-size: 22px;

    border-left: solid 7px #ffaf58;
  }
</style>