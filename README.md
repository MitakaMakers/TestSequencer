[日本語](README.ja.md) | [English](README.md)

# What is Excel Command?
Excel Command is free software for easy control of measuring instruments. 
Communication commands can be written in an Excel file and sent and received by various communication interfaces. 
Our goal is to create a calibration automation software like Fluke's MET/CAL in the future.

# Features
- Works with ExcelVBA (spreadsheet macro language) and TMCTL.DLL. Any installation work is not required.
- Communicates with measuring instruments via GP-IB, RS232, USB, or LAN.
- Controls and measures up to 8 measuring instruments.
- Japanese and English text can be switched.
- The VBA part is released as open source software.

# Target devices
- Measuring instruments compliant with IEEE488.2-1989

# Operating environment
- Windows: XP, 7, 10, 11
- Excel: 2010, 2013, 2016, 2019, 2021
- Office for Mac and Microsoft 365 are not supported.

# Communication interface
- GP-IB: Environment in which National Instruments GP-IB interface works.　* Install NI-488.2 separately.
- RS232C: Environment in which serial port or virtual COM port operates.
- LAN: Socket communication, VXI-11, or HiSLIP must work.
- USB: Environment where National Instruments NI-VISA or Yokogawa Measurement USB driver works.　*Please install NI-VISA or Yokogawa USB driver separately.

## How to use
## Download and Extract
After downloading the ZIP file, extract the 5 file (ExcelCommand.xlsm, tmctl.dll, tmctl64.dll, YKMUSB.dll and YKMUSB64.dll), and place these files in the same directory.

## Open a book containing macros
By default Excel displays a security warning and disables macros when you try to open a book containing macros. To enable or disable macros, follow these steps

- In an Excel program, on the File tab, click Options.
- Click Trust Center, and then click Trust Center Settings. The following screen is an example from
- Click "Macro Settings".
- Select "Disable VBA macros except digitally signed macros" from the "Macro Settings" list

## Address string
(under construction)

## Instruction
(under construction)

## Copyright statement as open source software
Excel Commmand: An excel macro file to communicate some measurement insturuments.
Copyright (C) 2023 Takatoshi Yamaoka

This program is free software: you can redistribute it and/or modify
This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as
This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as
License, or any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
If not, see <https://www.gnu.org/licenses/>.