# Invoice-Generater
Invoice generater base on Google Drive.

Code.gs should be attached into data.gsheet by App Script

**file structure**
```
Folder
├── TEMP 
├── PDF
│   └── [Output PDF file]
├── Data.gsheet
└── Template.gdoc
```

 **data structure in Data.gsheet**

| Sheet | Column A | Column B | Column C | Column D | Column E | Column F |
| --- | --- | --- | --- | --- | --- | --- |
| `Test` | number | name | date | item | qty | total |
| `From` | myname | myabn | myphone | myaddress_1 | myaddress_2 |
| `To` | customername | customerabn | customerphone | customeraddress_1 | customeraddress_2 |

Sheet Generate:
![Imgur](https://i.imgur.com/tkBMQ2U.png)


**invoice template**
  
![Imgur](https://i.imgur.com/04gycqj.png)

