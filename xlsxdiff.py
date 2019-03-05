import pandas as pd
import numpy as np
import click

# idea from:
# https://kanoki.org/2019/02/26/compare-two-excel-files-for-difference-using-python/


@click.command()
@click.option('--file1', required=True, help='path to 1st xlsx file')
@click.option('--file2', required=True, help='path to 2nd xlsx file')
def main(file1, file2):
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
    except Exception as e:
        print(f'Encountered Exception: {e} '
              f'while reading files to DataFrame, '
              f'file1: {file1}, file2: {file2}')
        return

    try:
        df1 = df1.fillna('')
        df2 = df2.fillna('')
    except Exception as e:
        print(f'failed to strip nan from dataframes with exception : {e}')
        return

    try:
        df1.equals(df2)
    except Exception as e:
        print(f'{file1} and {file2} do not contain matching columns'
              'and cannot be diff\'d, (see dataframe.equals docs)')
        return

    try:
        comparison_values = df1.values == df2.values
    except Exception as e:
        print(f'Failed to compare dataframes with exception: {e}')
        return

    try:
        rows, cols = np.where(comparison_values==False)
    except Exception as e:
        print(f'Failed to generate Rows and Columns fm comparison_values '
              f'with Exception: {e}')
        return

    try:
        for item in zip(rows, cols):
            df1.iloc[item[0], item[1]] = \
                f'{df1.iloc[item[0], item[1]]} --> ' \
                f'{df2.iloc[item[0], item[1]]}'
    except Exception as e:
        print(f'Failed to rewrite df1 rows with exception: {e}')
        return

    try:
        # TODO: better output file (not ./)
        output_file = './Excel_diff.xlsx'
        df1.to_excel(output_file, index=False, header=True)
        print(f'Output resulting diff to: {output_file}')
    except Exception as e:
        # TODO: add output filename to exception logging
        print(f'Failed to write to Excel file with exception: {e}')
        return

    return


if __name__ == '__main__':
    main()
