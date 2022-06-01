import pandas as pd
import numpy as np
import os

class Roller:
    """ A class to calculate a set of rolling functions
        on a pandas dataframe and to optionally clear cells
        between the specified interval
        
        Parameters:
        ===========
        functions: A list of function for aggregation
        col2agg: A string specifying the column to aggregate
        interval: An integer specify the rolling interval
                 to use for aggregation
        insheet: A string specifying the name of sheetname in `infile`
                 to be read
        insheet: A string specifying the name of sheetname in `outfile`
                 to be written 
        infile: A string of the path to the input Excel file
        outfile: A string of the path to the output Excel file
        ispickle: A boolean, is the `infile` a pickle file? 
        **kwargs: Any number of keyword arguements that can be passed to 
                  pandas.read_excel()
        
    """
    def __init__(self, functions: list, col2agg: str, interval: int,
             insheet: str, outsheet: str, infile: str, outfile: str,
             ispickle: bool, **kwargs) -> "Roller":
        self.functions = functions
        self.col2agg = col2agg
        self.interval = interval
        self.insheet = insheet
        self.outsheet = outsheet
        self.infile = infile
        self.outfile = outfile
        self.colnames = [f"avg-{self.interval}", f"std-{self.interval}"]
        # --- Read in the xlsx file as a pandas dataframe with choice columns --- #
        # Read a particular sheet from the excel file
        self.df =  pd.read_pickle(self.infile) if ispickle else pd.read_excel(self.infile, self.insheet,**kwargs)
        # Save  as pickle file if pickle file doesn't
        # Already exist
        if not ispickle:
            self.df.to_pickle(f"{self.outsheet}.pickle")
            
    @property
    def functions(self) -> list:
        return self._functions
    
    @functions.setter
    def functions(self, functions: list):
        try:
            assert isinstance(functions, list)
        except AssertionError:
            raise TypeError("functions must be of type list")
        if not functions:
            raise ValueError("You must povide a list of functions")
        self._functions = functions

    @property
    def col2agg(self) -> str:
        return self._col2agg 
    
    @col2agg.setter
    def col2agg(self, column: str):
        try:
            assert isinstance(column, str)
        except AssertionError:
            raise TypeError("Column must be of type str")
        if not column:
            raise ValueError("Column cannot be empty")
        self._col2agg = column
        
    # Create a function that will get the list of row indices
    # to clear from the dataframes
    def get_indices(self) -> list:
        """ A function to get the list of row indices to clear 
            in a dataframe
        """
        length = len(self.df)
        indices = np.arange(length+1)
        to_remove = [self.interval-1]
        right = list(range(to_remove[-1],length,self.interval))
        to_remove.extend(right)
        indices_to_remove = [np.where(indices == i) for i in to_remove]
        # Here's the game changer, used numpy's ufunc instead of a for loop
        self.indices = list(np.delete(indices, indices_to_remove))
        
        # If the last index is greater than the 
        # the available indices rcolnames it
        if self.indices[-1] >= length:
            self.indices.pop()

        return self.indices 
    
    # Calculate the rolling means and standard deviations 
    def aggregate(self):
        """A function to calculate a rolling set of funcions
           on a pandas dataframe.
        """
        self.df[self.colnames] =  (self.df.rolling(self.interval)[self.col2agg]
                                .agg(self.functions))
    # Clear unwater cells    
    def clear_cells(self):
        """ A function to clear unawanted cells given a set 
            of indices to clear from the dataframe.
        """
        df = self.df.copy()
        df.loc[self.indices,self.colnames] = ''
        return df
    
    # Write to excel    
    def write2excel(self):
        if os.path.exists(self.outfile):
            with pd.ExcelWriter(self.outfile, engine='openpyxl', mode='a') as writer:
                self.clear_cells().to_excel(writer, sheet_name=self.outsheet, index=False)
        else:
            with pd.ExcelWriter(self.outfile, engine='openpyxl') as writer:
                self.clear_cells().to_excel(writer, sheet_name=self.outsheet, index=False)
                
    def run(self, write2excel: bool = False):
        """ A method to run the complete pipeline
        
            Parameters:
            ===========
            write2excel: A boolean that if set to True will write 
                       the modified dataframe to Excel
        """
        self.get_indices()
        self.aggregate()
        if write2excel:
            self.write2excel()
        else:
            return self.clear_cells()
 
    def __repr__(self):
        return (f"Roller({self.functions}, {self.col2agg}, {self.interval}, "
                f"{self.insheet}, {self.outsheet},{self.infile}, {self.outfile}, **kwargs)")
    
    @classmethod
    def get(cls):
        pass
        #return cls(functions, column)