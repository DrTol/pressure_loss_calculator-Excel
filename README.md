# Pressure Loss Calculator Tool - Excel (SI Units)
An Excel add-in with User-Defined-Functions (VBA) to calculate the friction pressure loss (head loss) in circular pipes with full flow water. The pressure loss calculator function has the selective options for the friction loss formulations, can be selected as either "Hazen-Williams" or as "Darcy-Weisbach". For the latter, another selective options are possible for the equation calculating of the Darcy-Weisbach friction factor (details given in [Features](README.md#features)). Besides, another functions are given to convert between the roughness factors/coefficients, made use of in the Darcy-Weisbach and the Hazen-Williams formulations. The Matlab calculator tools of the same Excel functions can be found in [pressure_loss_calculator-Matlab @GitHub](https://github.com/DrTol/pressure_loss_calculator-Matlab.git).

## Table of Contents
- [Features](README.md#features)
- [Usage](README.md#usage)
  - [Description of the Example Excel File](README.md#description-of-the-example-excel-file)

## Features 
- The VBA user-defined functions are available as packed in the Excel add-in [ExcelAdd-In_PressureLoss.xlam](https://github.com/DrTol/pressure_loss_calculator-Excel/blob/master/ExcelAdd-In_PressureLoss.xlam) but also available on an individual basis for each function in the GitHub folder [Modules-UDFs](https://github.com/DrTol/pressure_loss_calculator-Excel/tree/master/Modules-UDFs)), the details for the functions described below: 
- The function for the calculation of the pressure loss has the options in selecting the solver type through the equations either by 'Darcy-Weisbach' or by 'Hazen-Williams'. 
- Besides, another feature in the tool options allows users to select through various algorithms to calculate the Darcy-Weisbach friction coefficient *f*, limited to algorithms by 'Moody', 'Colebrook-White', 'Clamond', 'Swamee-Jain', 'Zigrang-Sylvester', and 'Haaland'. 
- Aside from the pressure loss calculation function, two other converter tools are also given to obtain the Hazen-Williams roughness coefficient *C* as a non-steady variable by a function of (i) the absolute roughness of the pipe (also known as ε - eps) and (ii) the Darcy-Weisbach friction factor *f*.
- The limitations for use of equations and algorithms are given in the code (e.g. the operational limitations in using Hazen-Williams).

## Usage
An example Excel file is given in [examplePressureLoss&RoughnessConverters.xlsm](examplePressureLoss%26RoughnessConverters.xlsm) to illustrate the use of the user-defined-functions in question. 
WARNING: please do not load the developed Excel add-in [ExcelAdd-In_PressureLoss.xlam](ExcelAdd-In_PressureLoss.xlam) in this stand-alone example Excel file because it already involves of the functions packed in this Excel add-in. 

### Description of the Example Excel File
There are four different worksheets in this example Excel file [examplePressureLoss&RoughnessConverters.xlsm](examplePressureLoss%26RoughnessConverters.xlsm), each hosts to illustration of usage for different functions, details given below. It should be noted that each Excel WorkSheet has the layout of *INPUTS* on the left-hand side and the *USAGE EXAMPLE* on the right-hand side, latter involves of the calculator functions in use. Another note is that this Excel file follows a consistent formatting by use of cell styles (e.g. calculation results in *calculation* style, the results in *output* style etc.). 

- **WorkSheet _"PressureLossCalculator"_:** This first Excel WorkSheet *PressureLossCalculator* shows the usage of the pressure loss calculator function *PressureLoss(L, D, aRou, mFlow, T, P, Solver, Algorithm, fTol, MaxIter)*, respectively, the input arguments being the length of the pipe segment, the water mass flow rate, the water temperature, the hydrostatical water pressure, the solver as optional (the default is as "Darcy-Weisbach"), the algorithm  as optional (the default is as "Clamond"), and iteration inputs as optional (valid only for the algorithm "Colebrook-White") fTol as the iteration tolerance and MaxIter as the maximum amount of steps for the iteration. 
For example, the cell "G5" host the pressure loss calculation as based on the "Darcy-Weisbach" solver with "Clamond" algorithm and the cell "G14" as based on the "Hazen-Williams" solver (no friction factor algorithm this time - please trace the precendents from the INPUTS section). 

- **WorkSheet _"DarcyFrictionAlgorithms"_ :** This Excel WorkSheet illustrates the usage of the Darcy-Weisbach friction factor algorithms. 
For example, the cell "F5" has the return from the [f_ColebrookWhite function](Modules-UDFs/dwf_ColebrookWhite.bas) (without iteration input required in do-while loop), the cell "F6" has the return from the [f_Moody function](Modules-UDFs/dwf_Moody.bas) etc. 

- **WorkSheet _"Converter_f2C"_ :** This Excel WorkSheet illustrates the usage of the converter function [tConverterDW2HW](Modules-UDFs/Converter_f2C.bas) that is developed to convert the Darcy-Weisbach friction factor *f* to Hazen-Williams roughness coefficient *C* and the vice versa conversion from *C* to *f*. The two examples are that i) "H3" cell is the return as *f* as converted from the *C* given in "F3" and ii) "H6" cell is the return as *C* as converted from the *f* given in "F6"  

- **WorkSheet _"Converter_Roughness"_ :** This Excel WorkSheet illustrates the usage of the converter function [tConverterRoughness](Modules-UDFs/Converter_rRou2C.bas) that is developed to convert the relative roughness (absolute pipe roughness / pipe diameter - *eps/D* or *rRou*) to Hazen-Williams roughness coefficient *C* and vice versa. The two examples are that i) "H3" cell is the return as *C* as converted from the *rRou* given in "F3" and ii) "H7" cell is the return as *rRou* as converted from the *C* given in "F7" 

- **Other WorkSheets :** The other Excel WorkSheets *"zPipeCatalogue"* and *"zDiskinData"* are not an illustration of usage but a required data for this example Excel file [examplePressureLoss&RoughnessConverters.xlsm](examplePressureLoss%26RoughnessConverters.xlsm). The *zDiskinData* is necessary for the user-defined-functions developed (given in this repository), i.e. especially at functions [tConverterRoughness](Modules-UDFs/Converter_rRou2C.bas) and [tReynoldsLimits](Modules-UDFs/tHWLimitsReynolds.bas).
