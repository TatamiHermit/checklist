Checklist
==========
Using Pandas & xlwings.  
For infotainment test report data cleaning.  

Notice: 
==========
# 1.appending libs: pandas xlwings xlrd  
Terminal:  
```python  
pip install pandas  
pip install xlwings  
pip install xlrd  
```  
# 2.put xlsx and python file in the same directory  
# 3.run checklist.py  
---
OR   
Use pyinstaller package checklist.py to exe.  
terminal:
```python 
pip install pyinstaller  
``` 
terminal:  
```python
pyinstaller -F $FileNamePath -i $FileIcoPath  
``` 


ReleaseNote：  
==========
20190819
---
fix bug name_check kill excel progress  

20190816  
---
1 modify sumlist  
2 add step state check  
3 add step overall state check  

20190805  
---
1 initial release   
