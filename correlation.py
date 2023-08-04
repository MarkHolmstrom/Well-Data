#!/usr/bin/env python3

import sys

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog


def get_regression_coefficients(df):
    X = np.vstack((np.multiply(df["DEPT"], np.square(df["DTCO"])), np.square(df["DTCO"]), df["DTCO"])).T
    # regression coefficient matrix
    least_squares_system = np.linalg.lstsq(X, df["DTSM"], rcond=None)  # predicting DTSM
    # 0 is DTCO coefficient, 1 is depth coefficient
    coefficient_vector = least_squares_system[0]
    sum_of_squared_residuals = least_squares_system[1][0]
    plt.scatter(df["DTCO"], df["DTSM"], s=1, c=df["DEPT"])
    x_plot = np.linspace(0, df["DTCO"].max(), 100)
    y_plot = coefficient_vector[1] * x_plot ** 2 + coefficient_vector[2] * x_plot  # does not account for interaction
    plt.plot(x_plot, y_plot, c="tab:orange")
    plt.xlabel("DTCO (us/m)")
    plt.ylabel("DTSM (us/m)")
    colorbar = plt.colorbar(label="DEPTH (m)")
    colorbar.ax.invert_yaxis()  # more sensible display
    plt.show()

    coefficient_df = pd.DataFrame([coefficient_vector], columns=["DTCO^2*DEPT", "DTCO^2", "DTCO"])

    total_sum_of_squares = np.sum(np.square(np.mean(df["DTSM"]) - df["DTSM"]))
    coefficient_determination = 1 - sum_of_squared_residuals / total_sum_of_squares
    RMSE = (sum_of_squared_residuals / len(df)) ** 0.5
    MAPE = np.mean(np.divide(abs(np.dot(X, coefficient_vector) - df["DTSM"]), df["DTSM"]))
    return coefficient_df, coefficient_determination, RMSE, MAPE


# noinspection PyTypeChecker
def get_dtco_dtsm_dataframe(path_list):
    df = pd.DataFrame(columns=["DTCO", "DTSM"])
    for path in path_list:
        try:
            log_df = pd.read_excel(path,
                                   sheet_name="logs", header=1, skiprows=[2], usecols=["DEPT", "DTCO", "DTSM"])
        except ValueError:  # 100060307913W600 log uses different format (note two of each col, m is before ft so used)
            log_df = pd.read_excel(path,
                                   sheet_name="mechanics", header=9, skiprows=[10], usecols=["DEPT", "DTCO", "DTSM"])
        log_df = log_df.dropna()  # dropping rows with missing values
        log_df = log_df[(log_df > 0).all(1)]  # dropping rows with negative values
        log_df = log_df.drop_duplicates(subset=["DTCO", "DTSM"])  # dropping rows with repeated wave measurement data
        df = pd.concat([df, log_df])
        print("Processed", path)
    df = df.reset_index(drop=True)  # reset index
    return df


def main():
    app = QApplication(sys.argv)  # just to keep QApplication in memory, a gui event loop with exec_() isn't needed
    print("Select the well logs that have DTCO and DTSM data")
    log_paths = QFileDialog.getOpenFileNames(filter="All Files (*)")[0]
    if not log_paths:  # cancelled
        sys.exit()
    df = get_dtco_dtsm_dataframe(log_paths)
    coefficient_df, coefficient_determination, RMSE, MAPE = get_regression_coefficients(df)

    print("Predicting DTSM as a linear combination of DTCO and Depth terms:\n",
          coefficient_df.to_string(index=False),
          "\nCoefficient of determination (r^2): ", coefficient_determination,
          "\nRoot Mean Square Error: ", RMSE,
          "\nMean Absolute Percentage Error: ", MAPE)

    save_file = QFileDialog.getSaveFileName(filter="CSV Files (*.csv);;All Files (*)")[0]
    if not save_file:  # cancelled
        sys.exit()
    coefficient_df[["r^2", "RMSE", "MAPE"]] = coefficient_determination, RMSE, MAPE
    coefficient_df.to_csv(save_file, index=False)


if __name__ == "__main__":
    main()
