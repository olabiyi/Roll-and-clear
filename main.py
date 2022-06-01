import pandas as pd
import numpy as np
import os
from roller import Roller
from argparse import ArgumentParser





def main():
    # ------- Parse commandline arguements and Set all necessary variables ---------------- #
    # Working from within the directory where Roller.py is located
    parser = ArgumentParser()

    parser.add_argument("-i","--infile", help="Input excel or pickle file name", type=str)
    parser.add_argument("-o","--outfile", help="Output excel file name", type=str, default="new.xlsx")
    parser.add_argument("-c","--colnames", help="Comma separated list of column names in the input File", type=str)
    parser.add_argument("-a","--column2aggregate", help="Column names in the input Excel File", type=str)
    parser.add_argument("-f","--aggregate_functions", help="Comma delimited list of 2 functions to apply", 
                        default="mean,std", type=str)
    parser.add_argument("-r","--rolling_interval", help="Rolling interval to use for aggregation", default=5, type=int)
    parser.add_argument("-s","--insheet", help="Rolling interval to use for aggregation", type=str)


    args = parser.parse_args()

    col_list = [col for col in args.colnames.split(",")]
    aggregate_functions = [function for function in args.aggregate_functions.split(",")]
    outsheeet =  f"edited-{args.insheet}"

    # Check to see if the input file is a pickle or excel file
    ispickle=False if args.infile.endswith("xlsx") else True
    # --- Create a Roller object#
    df = Roller(aggregate_functions, args.column2aggregate, args.rolling_interval,
              args.insheet, outsheeet, args.infile, args.outfile,
              ispickle, header=1, usecols=col_list)

    # Let roll baby
    df.run(True)
    
if __name__ == "__main__":
    main()