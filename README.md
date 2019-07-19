# UHV_chamber_auto_design

This is a VBA code that automatically builds a UHV chamber 3D CAD model using Solidworks API. Parameters defining port geometry are contained in the code. The code was tested and verified with Solidworks 2017 version on 10 Feb 2018. The generation process should proceed like [this video](https://www.youtube.com/watch?v=t8tN-IW61VU). Refer to the [MDC article](https://www.mdcvacuum.com/displayContentPageFull.aspx?cc=CUSTOMENG) for convention on chamber geometry. Although written in VBA, this code should help users of other programming languagues resolve various technical issues that are not covered or described incorrectly in the Solidworks API manual, and that one can not figure out using Solidworks macro. This code would salvage API-based chamber designers by making them not wast time and effort on the troubles I had to go through.

## Prerequisite for code execution

1. Add references for sldworks and swconst into your project. They were located in C:\Program Files\SolidWorks Corp\SolidWorks\api\redist\Solidworks.Interop.sldworks.dll and C:\Program Files\SolidWorks Corp\SolidWorks\api\redist\Solidworks.Interop.swconst.dll in my case.
2. Update 'rootFolder' variable in line 72 to your root folder
3. Check, and update as necessary, the locations and the file names of 'part.prtdot' (line 151) / assem.asmdot (line 226). They may vary depending on the Solidworks version.
4. Download flange files to your local PC into 'Flanges' folder under the root folder. Or you can update line 256 to your folder.
5. Now execute the code.

## Notes

1. Flange files were all drawn from scratch by me to MDC CAD models and specs.
2. CF2750 and CF4500 ports have bigger-than-standard diameter. You may want to update the parts file and diameter in the source code.
3. The present source code uses not all of the uploaded flanges. But all flange files I built were uploaded to make this project available for broader application.
4. I recommend you not to disturb chamber generation process by clicking any part of Solidworks window or any other part of your operating system. Your action may interfere with the process.
5. Fundamental reference planes are referred to as "Front Plane" in Solidworks 2014 but simply "Front" in 2017. Check yours and properly update plane designators wherever necessary. For instance lines 242/243, 246/247, 250/251, 272, 278, 282, etc. Flange parts files were made with Solidworks 2014. That is the reason for the mixed nomenclatures.
