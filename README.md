# UHV_chamber_auto_design

This is a VB.net code that automatically builds a UHV chamber 3D CAD model using Solidworks API. Parameters defining port geometry are contained in the code. The code was tested and verified with Solidworks 2017 version on 10 Feb 2018. The generation process should proceed like [this video](https://www.youtube.com/watch?v=iZ9ieSHOZsg). Refer to the [MDC article](https://www.mdcvacuum.com/displayContentPageFull.aspx?cc=CUSTOMENG) for convention on chamber geometry. Although written in VB.net, this code should help users of other programming languagues resolve various technical issues that are not covered or described incorrectly in the Solidworks API manual, and that one can not figure out using Solidworks macro. This code hopefully would salvage API-based chamber designers by making them not wast time and effort on the troubles I had to go through.

## Prerequisite for code execution

1. Add references for sldworks / swconst into your project. They are located in C:\Program Files\SolidWorks Corp\SolidWorks\api\redist\sldworks.dll and C:\Program Files\SolidWorks Corp\SolidWorks\api\redist\swconst.dll in my case.
2. Update 'rootFolder' variable in line 72 to your root folder
3. Check, and update as necessary, the locations and the file names of 'part.prtdot' (line 151) / assem.asmdot (line 226). They may vary depending on the Solidworks version.
4. Download flange files to your local PC into 'Flanges' folder under the root folder. Or you can update line 256 to your folder.
5. Now execute the code.

## Notes
