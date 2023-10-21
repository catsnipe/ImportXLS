This is an older converter.
Newer converters are registered here.

→ https://github.com/catsnipe/XlsToJson

# ImportXLS
Convert excel sheet to Unity Scriptable object

![image](https://user-images.githubusercontent.com/85425896/122719958-d3d21a80-d2a9-11eb-83b0-cda205fc1749.png)

## requirement
unity2017 or later  
npoi2.5.1 or later  

## usage
1. Download sample project and open `SampleScene.unity`.
2. Right click on Sample.xlsx and `ImportXLS`.
3. Click 'CREATE ALL'.  
   Source code is automatically generated.  
   
      * **Scripts/Global.cs**  
         defines an enum.
      * **Scripts/Sample.cs**  
         a table for accessing the table.  
      * **Scripts/X_Sample.cs**  
         a singleton class for easy handling of Sample Class.  
      * **Editor/ImportXLS/importer/Data_Sample_Import.cs**  
         a class for creating a Scriptable Object of the Sample class.  
   
4. Right click on Sample.xlsx and `Reimport`.  
   Scriptable Object is automatically generated.  
   
      * **Resources/Data_Sample.asset**  
         an Excel table whose type is defined in Sample Class.

5. Modify 'DebugTableDraw.cs' and then run unity.  
   `#if TRUE -> #if FALSE`  

6. If the contents of the Sample sheet are displayed in the debug console, it is successful.
   ![image](https://user-images.githubusercontent.com/85425896/122723793-5230bb80-d2ae-11eb-9904-7d46632d8614.png)

7. After that, if you change the table, `Reimport` it.

see more detail (japanese): https://www.create-forever.games/xls-scriptable3/

## license
This sample project includes the work that is distributed in the Apache License 2.0 / MIT / MIT X11.  

NPOI (Apache2.0): https://www.nuget.org/packages/NPOI/2.5.1/License  
SharpZLib (MIT): https://licenses.nuget.org/MIT  
Portable.BouncyCastle (MIT X11): https://www.bouncycastle.org/csharp/licence.html  
