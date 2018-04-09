# Pressure Loss Calculator Tool - Excel (SI Units)
An Excel add-in with User-Defined-Functions (VBA) to calculate the friction pressure loss (head loss) in circular pipes with full flow water. The pressure loss calculator function has the selective options for the friction loss formulations, can be selected as either "Hazen-Williams" or as "Darcy-Weisbach". For the latter, another selective options are possible for the equation calculating of the Darcy-Weisbach friction factor (details given in [Features](README.md#features)). Besides, another functions are given to convert between the roughness factors/coefficients, made use of in the Darcy-Weisbach and the Hazen-Williams formulations. The Matlab calculator tools of the same Excel functions can be found in [pressure_loss_calculator-Matlab @GitHub](https://github.com/DrTol/pressure_loss_calculator-Matlab.git).

## Table of Contents
- [Features](README.md#features)
- [Usage](README.md#usage)
  - [Description of the Example Excel File](README.md#description-of-the-example-excel-file)
  - [Description of the Excel Add-In](README.md#description-of-the-excel-add-in)
  - [Description of Modules&Functions](README.md#description-of-modulesfunctions)

## Features 
- The VBA user-defined functions are available as packed in the Excel add-in [ExcelAdd-In_PressureLoss.xlam](https://github.com/DrTol/pressure_loss_calculator-Excel/blob/master/ExcelAdd-In_PressureLoss.xlam) but also available on an individual basis for each function in the GitHub folder [Modules-UDFs](https://github.com/DrTol/pressure_loss_calculator-Excel/tree/master/Modules-UDFs)), the details for the functions described below: 
- The function for the calculation of the pressure loss has the options in selecting the solver type through the equations either by 'Darcy-Weisbach' or by 'Hazen-Williams'. 
- Besides, another feature in the tool options allows users to select through various algorithms to calculate the Darcy-Weisbach friction coefficient *f*, limited to algorithms by 'Moody', 'Colebrook-White', 'Clamond', 'Swamee-Jain', 'Zigrang-Sylvester', and 'Haaland'. 
- Aside from the pressure loss calculation function, two other converter tools are also given to obtain the Hazen-Williams roughness coefficient *C* as a non-steady variable by a function of (i) the absolute roughness of the pipe (also known as ε - eps) and (ii) the Darcy-Weisbach friction factor *f*.
- The limitations for use of equations and algorithms are given in the code (e.g. the operational limitations in using Hazen-Williams).

## Usage
A stand-alone example Excel file is given in [examplePressureLoss&RoughnessConverters.xlsm](examplePressureLoss%26RoughnessConverters.xlsm) to illustrate the use of the user-defined-functions in question. 
**WARNING:** please do not load the developed Excel add-in [ExcelAdd-In_PressureLoss.xlam](ExcelAdd-In_PressureLoss.xlam) in this stand-alone example Excel file because it already involves of the functions packed in this Excel add-in. Otherwise you will have two of the same functions and modules in your VBA library, one from the example Excel file and the other from the Excel add-in!!!

### Description of the Example Excel File
There are four different worksheets in this example Excel file [examplePressureLoss&RoughnessConverters.xlsm](examplePressureLoss%26RoughnessConverters.xlsm), each hosts to illustration of usage for different functions, details given below. It should be noted that each Excel WorkSheet has the layout of *INPUTS* on the left-hand side and the *USAGE EXAMPLE* on the right-hand side, latter involves of the calculator functions in use (Check the [ScreenShot](zScreenShots/ScreenShot_How2Use_PressureLossCalculator.png)). Another note is that this Excel file follows a consistent formatting by use of cell styles (e.g. calculation results in *calculation* style, the results in *output* style etc.). 

- **WorkSheet _"PressureLossCalculator"_:** This first Excel WorkSheet *PressureLossCalculator* shows the usage of the pressure loss calculator function [PressureLossCalculator.bas](Modules-UDFs/PressureLossCalculator.bas) as *PressureLoss(L, D, aRou, mFlow, T, P, Solver, Algorithm, fTol, MaxIter)*, respectively, the input arguments being the length of the pipe segment, the water mass flow rate, the water temperature, the hydrostatical water pressure, the solver as optional (the default is as "Darcy-Weisbach"), the algorithm  as optional (the default is as "Clamond"), and iteration inputs as optional (valid only for the algorithm "Colebrook-White") fTol as the iteration tolerance and MaxIter as the maximum amount of steps for the iteration.
For example, the cell "G5" host the pressure loss calculation as based on the "Darcy-Weisbach" solver with "Clamond" algorithm and the cell "G14" as based on the "Hazen-Williams" solver (no friction factor algorithm this time - please trace the precendents from the INPUTS section). 

- **WorkSheet _"DarcyFrictionAlgorithms"_ :** This Excel WorkSheet illustrates the usage of the Darcy-Weisbach friction factor algorithms. 
For example, the cell "F5" has the return from the [f_ColebrookWhite function](Modules-UDFs/dwf_ColebrookWhite.bas) (without iteration input required in do-while loop), the cell "F6" has the return from the [f_Moody function](Modules-UDFs/dwf_Moody.bas) etc. 

- **WorkSheet _"Converter_f2C"_ :** This Excel WorkSheet illustrates the usage of the converter function [tConverterDW2HW](Modules-UDFs/Converter_f2C.bas) that is developed to convert the Darcy-Weisbach friction factor *f* to Hazen-Williams roughness coefficient *C* and the vice versa conversion from *C* to *f*. The two examples are that i) "H3" cell is the return as *f* as converted from the *C* given in "F3" and ii) "H6" cell is the return as *C* as converted from the *f* given in "F6"  

- **WorkSheet _"Converter_Roughness"_ :** This Excel WorkSheet illustrates the usage of the converter function [tConverterRoughness](Modules-UDFs/Converter_rRou2C.bas) that is developed to convert the relative roughness (absolute pipe roughness / pipe diameter - *eps/D* or *rRou*) to Hazen-Williams roughness coefficient *C* and vice versa. The two examples are that i) "H3" cell is the return as *C* as converted from the *rRou* given in "F3" and ii) "H7" cell is the return as *rRou* as converted from the *C* given in "F7" 

- **Other WorkSheets :** The other Excel WorkSheets *"zPipeCatalogue"* and *"zDiskinData"* are not an illustration of usage but a required data for this example Excel file [examplePressureLoss&RoughnessConverters.xlsm](examplePressureLoss%26RoughnessConverters.xlsm). The *zDiskinData* is necessary for the user-defined-functions developed (given in this repository), i.e. especially at functions [tConverterRoughness](Modules-UDFs/Converter_rRou2C.bas) and [tReynoldsLimits](Modules-UDFs/tHWLimitsReynolds.bas).

### Description of the Excel Add-In
The Excel add-in [ExcelAdd-In_PressureLoss.xlam](ExcelAdd-In_PressureLoss.xlam) allows using of the developed user-defined-functions in any Excel file that your calculations take part. How to install the Excel add-in is well described in [Acompara J - How to Install an Excel Add-In - Guide @ExcelCampus.com](https://www.excelcampus.com/tools/how-to-install-an-excel-add-in-guide/).

### Description of Modules&Functions
All of the modules developed within this repository are given in the folder [Modules-UDFs](Modules-UDFs). Generally, each of these modules hosts to a unique Excel functions developed, as shown in the table at section [List of Functions](README.md#list-of-functions). Here the idea is to present an overview of the Excel functions. Besides, one can use some of the Excel functions developed partially if other functions are not required by simply importing the bas file/s or copy&pasting them on the Visiual Basic Editor. 

Please note that, after copy&paste, you have to delete the first codeline in the .bas file. For example, if you need only of the converter function, after copy&paste, you have to delete the line _Attribute VB_Name = "Converter_f2C"_ in the [Converter_f2C.bas](Modules-UDFs/Converter_f2C.bas) or (another example) if you need only of the Clamond algorithm you have the delete the _Attribute VB_Name = "DWf_Clamond"_ from the code lines of [dwf_Clamond.bas](Modules-UDFs/dwf_Clamond.bas). [A screenshot of How2Do!](zScreenShots/DeleteAttributes-f_Clamond.png). 

### List of Functions
Here is the list of functions developed and in use (latter original works by other Developers):

| Function | Description | Module |
| --- | --- | --- | 
| **PressureLossCalculator** (L, D, aRou, mFlow, T, P, Solver, Algorithm, fTol, MaxIter) | The main function calculating the pressure loss | [PressureLossCalculator.bas](Modules-UDFs/PressureLossCalculator.bas) |
| **tConverterDW2HW** (f_or_C, D, Re, T, P) | The converter tool among the Darcy friction factor _f_ and the Hazen-Williams roughness coefficient _C_ | [Converter_f2C.bas](Modules-UDFs/Converter_f2C.bas) |
| **tConverterRoughness** (rRou_or_C, ConverDir) | The converter tool among the relative roughness _rRou_ and the Hazen-Williams roughness coefficient _C_ | [Converter_rRou2C.bas](Modules-UDFs/Converter_rRou2C.bas) |
| **tReynoldsLimits** (rRou_or_C, InputType) | Returns the limitations for the Reynolds range applicable for a given relative roughness or _C_ value | [tHWLimitsReynolds.bas](Modules-UDFs/tHWLimitsReynolds.bas) |
| **f_ColebrookWhite** (D, Re, rRou, fTol, MaxIter) | returns the Darcy-Weisbach friction factor by solving iteratively the Colebrook-White equation | [dwf_ColebrookWhite.bas](Modules-UDFs/dwf_ColebrookWhite.bas) |
| **f_Clamond** (Re, rRou) | Function returning the Darcy-Weisbach friction factor by use of the Clamond algorithm | [dwf_Clamond.bas](Modules-UDFs/dwf_Clamond.bas) |
| **f_Moody** (D, Re, aRou) | Function returning the Darcy-Weisbach friction factor by use of the Moody algorithm | [dwf_Moody.bas](Modules-UDFs/dwf_Moody.bas) |
| **f_Haaland** (D, Re, aRou) | Function returning the Darcy-Weisbach friction factor by use of the Haaland algorithm | [dwf_Haaland.bas](Modules-UDFs/dwf_Haaland.bas) |
| **f_SwameeJain** (D, Re, aRou) | Function returning the Darcy-Weisbach friction factor by use of the Swamee & Jain algorithm | [dwf_SwameeJain.bas](Modules-UDFs/dwf_SwameeJain.bas) |
| **f_ZigrangSylvester** (D, Re, aRou) | Function returning the Darcy-Weisbach friction factor by use of the Zigrang & Sylvester algorithm | [dwf_ZigrangSylvester.bas](Modules-UDFs/dwf_ZigrangSylvester.bas) |
| **Reynolds** (mFlow, D, T) | Function returning the Reynolds number as a function of the water mass flow rate | [tReynolds.bas](pressure_loss_calculator-Excel/Modules-UDFs/tReynolds.bas) |
| **LogBase** (x, base) | User-Defined VBA function returning the logarithm of a given number _x_ at a given base of _base_ | [zOtherTools.bas](Modules-UDFs/zOtherTools.bas) |
| **PiNumber** () | User-Defined VBA function returning the Pi number at the highest precision | [zOtherTools.bas](Modules-UDFs/zOtherTools.bas) |
| **Linterp** (KnownYs, KnownXs, NewX) | Linear interpolation function - developed by Wells, Ryan | [zInterp_Wells.bas](Modules-UDFs/zInterp_Wells.bas) |
| **XSteam Module** | A collection of functions returning the water properties at a given operational condition (e.g. _rhoL_T(T)_ returns the water density as a function of temperature _T_) - developed by Holmgren, Magnus | [zXSteam.bas](Modules-UDFs/zXSteam.bas) | 

## License
You are free to use, modify and distribute the code as long as authorship is properly acknowledged. The same applies for the original works 'XSteam' by Holmgren M. and 'colebrook.vb' by Clamond D, this repository functions make use of.

## Acknowledgement
We would like to acknowledge all of the open-source minds in general for their willing of share (as apps or comments/answers in forums), which has encouraged our department to publish the user-defined-functions developed during the PhD study here in GitHub.

This repository *pressure_loss_calculator-Excel* makes use of other original open-source projects: 
- [XSteam by Holmgren M.](http://xsteam.sourceforge.net/) | Author Description: XSteam provides accurate steam and water properties from 0 - 1000 bar and from 0 - 2000 deg C according to the standard IAPWS IF-97. For accuracy of the functions in different regions see IF-97 (www.iapws.org).
- [colebrook.m by Clamond D.](https://nl.mathworks.com/matlabcentral/fileexchange/21990-colebrook-m?focused=5105324&tab=function) - re-arranged as to the VBA programming language | Author Description: fast, accurate and robust computation of the Darcy-Weisbach friction factor _f_ according to the Colebrook equation. 
- [Function Linterp by Wells R.](https://wellsr.com/vba/2016/excel/powerful-excel-linear-interpolation-function-vba/) | Author Description: A simple and powerful Excel linear interpolation function.

The Excel Add-In (the user-defined VBA functions developed) as well as the Matlab tool of the same (released in [pressure_loss_calculator-Matlab.git](https://github.com/DrTol/pressure_loss_calculator-Matlab.git)) are by-products from the PhD study about the 4th generation (4G) low-temperature district heating systems in supply to low-energy houses, carried out by Hakan İbrahim Tol, PhD under the supervision of Prof. Dr. Svend Svendsen and Ass. Prof. Susanne Balslev Nielsen at the Technical University of Denmark (DTU). The PhD topic: "District heating in areas with low energy houses - Detailed analysis of district heating systems based on low temperature operation and use of renewable energy" - [free download by DTU (link)](http://orbit.dtu.dk/en/publications/district-heating-in-areas-with-low-energy-houses(9c056db5-8e76-425f-92ca-c072b642b6b3).html) or [by ResearchGate (link)](https://www.researchgate.net/publication/276266953_District_heating_in_areas_with_low_energy_houses_-_Detailed_analysis_of_district_heating_systems_based_on_low_temperature_operation_and_use_of_renewable_energy).

