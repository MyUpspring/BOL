# BOL Scripts

This script, `bol-walamrt.py`, is used to generate Wal-Mart's Bill of Lading (BOL).

This script, `bol-target.py`, is used to generate Target's Bill of Lading (BOL).

## Template
You don't need to touch the files under the [template](./template/) folder. 
[template-walmart.xlsx](./template/template-walmart.xlsx) is the BOL template for Wal-Mart. 
[template-target.xlsx](./template/template-target.xlsx) is the BOL template for Target 

## Usage

To use this script, follow these steps:
### prepare for the data
1. You should receive an email daily with the name `searchResult.csv` as the attachement. Download it and save it to [open_sales_orders.csv](./daily-data/open_sales_orders.csv)
2. You should get [routing_status.csv](./daily-data/routing_status.csv)
3. You should get [item_list.csv](./daily-data/item_list.csv) from NetSuite.com

Have those 3 files refreshed every time you wan to generate the BOL. `item_list.csv` might not change as often as other two data sources.

then you run this command in Powershell (enter `powershell` in your folder directory path) by typing:
```generate wal-mart's BOL
python bol-walmart.py
```

and, for Target
```generate Targt's BOL
python target.py
```

## Dependencies

This script requires the following dependencies:

- Python 3.9
- openpyxl 

## License

This project is licensed under the [License Name] license. See the [LICENSE](LICENSE) file for more details.

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## Contact

For any questions or inquiries, please contact [Liangjun Jiang] at [liangjun.jiang@upspringbaby.com].
