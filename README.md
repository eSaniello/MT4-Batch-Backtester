# MT4-Batch-Backtester
A simple script that backtests an EA on the pairs that I specified in one go.

## How it works:
- Open MT4
- Run EA on first pair in list
- Save report
- Close
- Repeat for all pairs in array

## Configuration in .vbs file
- data_dir_path = path to mt4 data folder
- base_dir_path = path to mt4 installation folder
- file_name = name of the .ini file
- expert = name of the expert advisor you want to test
- symbol_arr = array of pairs you want to run the backtest on
- specify start_date and end_date

## How to run:
- put the 01.ini file in the root of the mt4 data folder
- in mt4 go to settings and save a .set file for your ea and give it the same same as the .ini file
- make sure that the .ini and .set files are in the root of the datafolder
- run the .bat file
