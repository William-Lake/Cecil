from pathlib import Path

import pandas as pd


def build_df():
    
    df = None
    
    for xl_file in Path('.').glob('*.xlsx'):
        
        xl_df = pd.read_excel(xl_file,engine='openpyxl')
        
        xl_df['File'] = xl_file.name
        
        if df is None:
            
            df = xl_df
            
        else:
            
            df = pd.concat([df,xl_df],ignore_index=True)
            
    for col in df.columns:
        
        df[col] = df[col].apply(lambda val: str(val).upper())
            
    return df    

def main():
    
    df = build_df()
    
    while True:
        
        search_term = input('Search Term? >').strip()
        
        if not search_term: break
        
        results = {
            col:df[df[col].str.contains(search_term.upper())]
            for col
            in df.columns
        }
        
        results = dict(filter(
            lambda kv: not kv[1].empty,
            results.items()
        ))
        
        if results:
            
            num_cols = len(results)
            
            total_results = sum([
                len(rdf)
                for rdf
                in results.values()
            ])
            
            print(f'{total_results} results found in {num_cols} columns.',end='\n\n')
        
            for col, rdf in results.items():
                
                print(f'{len(rdf)} results found in column {col}',end='\n\n')
                
                for idx, row in rdf.iterrows():
                    
                    print('\t' + row[col],end='\n\n')
                    
                input('Press enter when ready to continue.\n\n')
        else:
            
            print('No results found.',end='\n\n')
            
if __name__ == '__main__':
    main()
    