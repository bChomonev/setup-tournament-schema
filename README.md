# Swiss and DE Tournament Simulation Scripts

## Overview

This repository contains two Python scripts designed to simulate different stages of a tournament: `simulate.py` and `dynamic.py`. Both scripts cater to organizing competitive events with a blend of fairness and competitiveness.

### `simulate.py`

`simulate.py` focuses on simulating a tournament using the Swiss system, ensuring fair matchups by pairing participants with similar records. It generates pairings, simulates rounds, and exports detailed results to an Excel file. Ideal for the initial stages of a tournament to determine top performers.

### `dynamic.py`

`dynamic.py` enhances tournament management by dynamically reading and updating standings from an Excel file and handling the Direct Elimination (DE) stage. It prepares for subsequent rounds based on ongoing results, simulates the DE matchups, and exports these to Excel, including the transition from Swiss to DE stages based on rankings.

## How to Use

1. Ensure Python and pandas are installed.
2. Customize participant lists and settings as needed.
3. Run `simulate.py` for the Swiss stage simulation.
4. Run `dynamic.py` to manage ongoing tournaments and handle the DE stage.

Refer to the enhanced "How to Use" section for setup and execution details. Both scripts offer a comprehensive solution for managing tournaments, from initial rounds to the climactic DE stage.

## Installation Instructions

Before running `simulate.py` and `dynamic.py`, ensure your environment is set up correctly.

### Prerequisites

- Python 3.x
- pandas library
- openpyxl library

### Step-by-step Guide

1. **Install Python:** If not already installed, download Python from the [official Python website](https://www.python.org/).

2. **Install Dependencies:** Open a terminal or command prompt. Install the required libraries by running:

   ```shell
   pip install pandas openpyxl

   ## Clone the Repository (if applicable)
   Clone or download the scripts from their repository to your local machine.

   ## Running the Scripts
   1. Navigate to the directory containing `simulate.py` and `dynamic.py`.
   2. To run `simulate.py`, execute the following command in your terminal:
   python simulate.py
   3. To run `dynamic.py`, execute the following command in your terminal:
   python dynamic.py
   ```

This setup ensures you have all necessary components to successfully simulate and manage tournament stages.
