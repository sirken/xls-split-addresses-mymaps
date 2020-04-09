# Split XLS, Retain Addresses
Splits an xls file with addresses for geocoding in Google MyMaps.

## Features

* Splits xls into 2000 rows by default ([MyMaps limit](https://support.google.com/mymaps/answer/3024836?co=GENIE.Platform%3DDesktop&hl=en))
* Retains mappable street addresses
* Removes any "PO Box" cells but retains City, State
* Removes lines that don't contain a valid address

## Installation

1. Install `openpyxl`
    ```bash
    pip3 install openpyxl
    ```

2. Set execute permissions for script
    ```bash
    chmod +x xls-addresses-split
    ```

3. Edit the script and specify the required header names that match your xls file
    ```bash
    xls_req_headers = [["Street",False,-1],["City",False,-1],["State",False,-1]]
    ```

## Usage

1. Split target xls into new files
    ```bash
    ./xls-addresses-split "file.xls"
    ```

2. You can optionally specify the number of rows
    ```bash
    ./xls-addresses-split "file.xls" 500
    ```
