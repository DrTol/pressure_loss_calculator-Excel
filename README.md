# Pressure Loss Calculator Tool - Excel (SI Units)
An Excel add-in with User-Defined-Functions (VBA) to calculate the friction pressure loss (head loss) in circular pipes with full flow water. The pressure loss calculator function has the selective options for the friction loss formulations, can be selected as either "Hazen-Williams" or as "Darcy-Weisbach". For the latter, another selective options are possible for the equation calculating of the Darcy-Weisbach friction factor (details given in [Features](https://github.com/DrTol/pressure_loss_calculator-Excel/blob/master/README.md#features)). Besides, another functions are given to convert between the roughness factors/coefficients, made use of in the Darcy-Weisbach and the Hazen-Williams formulations. The Matlab tools of the same calculator/s can be found in [pressure_loss_calculator-Matlab.git](https://github.com/DrTol/pressure_loss_calculator-Matlab.git).

## Table of Contents


## Features 
- The VBA user-defined functions are available as packed in the Excel add-in [ExcelAdd-In_PressureLoss.xlam](https://github.com/DrTol/pressure_loss_calculator-Excel/blob/master/ExcelAdd-In_PressureLoss.xlam) but also available on an individual basis for each function in the GitHub folder [Modules-UDFs](https://github.com/DrTol/pressure_loss_calculator-Excel/tree/master/Modules-UDFs)), the details for the functions described below: 
- The function for the calculation of the pressure loss has the options in selecting the solver type through the equations either by 'Darcy-Weisbach' or by 'Hazen-Williams'. 
- Besides, another feature in the tool options allows users to select through various algorithms to calculate the Darcy-Weisbach friction coefficient *f*, limited to algorithms by 'Moody', 'Colebrook-White', 'Clamond', 'Swamee-Jain', 'Zigrang-Sylvester', and 'Haaland'. 
- Aside from the pressure loss calculation function, two other converter tools are also given to obtain the Hazen-Williams roughness coefficient *C* as a non-steady variable by a function of (i) the absolute roughness of the pipe (also known as ε - eps) and (ii) the Darcy-Weisbach friction factor *f*.
- The limitations for use of equations and algorithms are given in the code (e.g. the operational limitations in using Hazen-Williams).

## Usage
An example Excel file is given in [examplePressureLoss&RoughnessConverters.xlsm](https://github.com/DrTol/pressure_loss_calculator-Excel/blob/master/examplePressureLoss%26RoughnessConverters.xlsm) to illustrate the use of the user-defined-functions in question. WARNING: please do not load the developed Excel add-in [ExcelAdd-In_PressureLoss.xlam](https://github.com/DrTol/pressure_loss_calculator-Excel/blob/master/ExcelAdd-In_PressureLoss.xlam) in this stand-alone example Excel file because it already involves of the functions packed in this Excel add-in. 

###### Description of the Example Excel File
There are four different worksheets in this example Excel file, each hosts to illustration of usage for different matters, details given below: 
- 

Please enjoy the example Excel file  making use of the modules/User-Defined-Functions(UDFs) (given in the folder [Modules-UDFs](https://github.com/DrTol/pressure_loss_calculator-Excel/tree/master/Modules-UDFs)) required for pressure loss calculation as built-in in this xlsm file - so you don't need to load the Excel Add-In  to this xlsm Excel file. 
