#!/usr/bin/env python3

import os
import sys

import lasio
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog


def merge_wells(df_list):
    """Given a list (ordered, mutable) of two dataframes containing well data, merge them by depth with the deeper
        well having priority in case of overlapping data."""
    if not df_list[0]["DEPT"].iat[0] < df_list[1]["DEPT"].iat[0]:
        df_list[0], df_list[1] = df_list[1], df_list[0]  # first (left) well is always shallower
    df = pd.merge(df_list[0], df_list[1], on="DEPT", how="outer").sort_values(by=["DEPT"])  # lexicographic fails
    order = [["DT_x", "DT_y"], ["GR_x", "GR_y"], ["RHOZ_x", "RHOZ_y"]]  # deeper well takes priority
    # order[0][0], order[0][1], order[1][0], order[1][1], order[2][0], order[2][1] = \
    #    order[0][1], order[0][0], order[1][1], order[1][0], order[2][1], order[2][0]
    if "DT_x" in df.columns:
        df["DT"] = df[order[0][1]].fillna(df[order[0][0]])
    if "GR_x" in df.columns:
        df["GR"] = df[order[1][1]].fillna(df[order[1][0]])
    if "RHOZ_x" in df.columns:
        df["RHOZ"] = df[order[2][1]].fillna(df[order[2][0]])
    return df[df.columns.intersection(["DEPT", "DT", "GR", "RHOZ"])]


def select_best_column(df, cols, pref=None):
    """Given a dataframe and a set of columns, return the column having the most valid entries.
        If none of the columns are in the dataset, i.e. the data is not in that file, return None.
        Missing data will be filled from other inputted columns if possible. Pref is the preferred column to use."""
    cols_adjusted = []
    for col in cols:
        for df_col in df.columns:
            if df_col == col or df_col.startswith(col + ":"):  # identically named columns in file
                cols_adjusted.append(df_col)
    if not cols_adjusted:
        return None  # data not in dataset
    else:
        if pref in df.columns:  # preferred column to use
            best_col = pref
        else:
            best_col = cols_adjusted[0]  # arbitrarily use the first one
        for col in cols_adjusted:  # fill in data from other columns
            try:
                df[best_col].fillna(df[col], inplace=True)  # in place modification
            except KeyError:
                pass
        return best_col


def process_well_folder(folder_path):
    df_list = []
    for file in os.listdir(folder_path):
        if file[-4:].lower() == ".las":  # alternatively, split basename to get file extension
            las = lasio.read(os.path.join(folder_path, file))
            df = las.df()
            df = df.reset_index()  # use depth as a column
            df[df < 0] = np.NaN  # all negative values considered invalid
            # df[df == las.well.null.value] = np.NaN
            df = df.dropna(subset="DEPT")  # drop rows with invalid (negative) depth e.g. DALMAIN.LAS
            if las.curves.DEPT.unit in {"F", "FT"}:  # data is in feet
                df["DEPT"] = df["DEPT"] * 0.3048
            dt_col = select_best_column(df, ["DT", "AC", "ACN"])
            gr_col = select_best_column(df, ["GR", "GRZ", "GRS", "GRD", "GRA"], pref="GRD")  # prefer GRD over GRS
            df = df[df.columns.intersection(["DEPT", dt_col, gr_col, "RHOZ"])]  # unique DEPT and RHOZ columns
            if dt_col in df.columns:
                df = df.rename(columns={dt_col: "DT"})
                if las.curves.DEPT.unit in {"F", "FT"}:  # data is in feet
                    df["DT"] = df["DT"] / 0.3048
            if gr_col in df.columns:
                df = df.rename(columns={gr_col: "GR"})
                # GR is metric/imperial agnostic
            if "RHOZ" in df.columns and las.curves.DEPT.unit in {"F", "FT"}:  # data is in feet
                df["RHOZ"] = df["RHOZ"] / (0.3048 ** 3)
            df_list.append(df)

    if len(df_list) == 2:  # priority logic implemented for two files
        # previous implementation
        """
        if df_list[0]["DEPT"].iat[0] < df_list[1]["DEPT"].iat[0]:  # first well is shallower
            df_list[1] = df_list[1][df_list[1]["DEPT"] > df_list[0]["DEPT"].iat[-1]]
        else:
            df_list[0] = df_list[0][df_list[0]["DEPT"] > df_list[1]["DEPT"].iat[-1]]
            df_list[0], df_list[1] = df_list[1], df_list[0]  # swapping dataframes is faster than resorting
        return pd.concat(df_list)
        """
        return merge_wells(df_list)
    else:  # single file
        return df_list[0]


def main():
    app = QApplication(sys.argv)  # just to keep QApplication in memory, a gui event loop with exec_() isn't needed
    print("Select the folder containing all north well folders to be processed, "
          "the output for each well will be saved as its folder name in the north well folder (csv formatted)")
    north_well_path = QFileDialog.getExistingDirectory()
    if not north_well_path:  # cancelled
        sys.exit()

    root, folder_names = next(os.walk(north_well_path))[:2]
    folder_paths = [os.path.join(root, folder) for folder in folder_names]
    for i, folder_path in enumerate(folder_paths):
        well_df = process_well_folder(folder_path)
        well_df = well_df[sorted(well_df.columns)]  # reorder in alphabetic order (matches given coincidentally)
        well_df.to_csv(os.path.join(root, folder_names[i]+".csv"), index=False)
        print("Processed", folder_path)
        # alternatively, split basename to get well name


if __name__ == "__main__":
    main()
