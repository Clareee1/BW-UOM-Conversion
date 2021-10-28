#!/usr/bin/env python
# coding: utf-8

# In[9]:


import pandas as pd
file_path = "UOM_Conversion_Input.xlsm"
df1 = pd.read_excel(file_path,'MainOffer')
#mainoffer
df1.head(10)


# In[10]:


import pandas as pd
import numpy as np
###---for BW one, always named after Price UOM---###
cus_uom,bw_uom = df1["Cust UOM"],df1["Price UOM"]

##iterate the column name
#for col in df1.columns:
 #   print(col)
    
###---"Packing""UoM""Price UOM" name might be different---###
df2 = df1[["Sort Key","Cust Description","Cust UOM","Part Description","Price UOM","Part Number"]]
##Assume all conversion is 1 currently
df2.loc[:, "conversion"] = 0
##all the Cust Pack Size & BW Pack Size = 1
df2.loc[:, "Cust_Size"] = 0
df2.loc[:, "BW_Size"] = 0
df2


# In[11]:


##convert from customer
## testing for customer uom
import re
for index, row in df2.iterrows():
    cust_uom = 1
    Bw_qty = 1
    ##customer size conversion
    ##customer description   
    if pd.isna(row['Cust UOM']):
        continue
        
    if re.match(r'(?i)(\d+)(m|cm|g|kg|ml|l|SH|SHT|Ltr|Lt|ltr|LTRS|GM)$',row['Cust UOM']):
        ###if BW UOM = each
        uom_match = re.match(r'(?i)(\d+)(m|cm|g|kg|ml|l|SH|SHT|Ltr|Lt|ltr|LTRS|GM)$',row['Cust UOM'])
        if row['Price UOM'] in (["EA","Each","ea","EACH"]):
            cust_value =  uom_match.group(1)
            cust_unit =  uom_match.group(2)
            bw_search = re.search(r'(?i)(\d+)(m|cm|g|kg|ml|l|SH|SHT|Ltr|Lt|ltr|LTRS|GM)$', row["Part Description"])
            if bw_search:
                bw_value = bw_search.group(1)
                bw_unit = bw_search.group(2)  
                if (cust_unit.lower() == bw_unit.lower()):
                    Bw_qty = bw_value
                    cust_uom = cust_value 
                    
                elif (len(cust_unit) >= len(bw_unit) and cust_unit[0:len(bw_unit)].lower() == bw_unit.lower()) :
                    Bw_qty = bw_value
                    cust_uom = cust_value
                    
                elif (len(cust_unit) < len(bw_unit) and bw_unit[0:len(cust_unit)].lower() == cust_unit.lower()):
                    Bw_qty = bw_value
                    cust_uom = cust_value
                    
                ## g vs kg || m vs cm || L(LTR) vs ML 
                elif (cust_unit.lower() == 'g' and bw_unit.lower() == 'kg') or (cust_unit.lower() == 'm' and bw_unit.lower() == 'cm') or (cust_unit.lower() in (["l","ltr"]) and bw_unit.lower() == 'ml'):
                    if (cust_unit.lower() == 'm' and bw_unit.lower() == 'cm'):
                        Bw_qty = bw_value/100  
                    else:
                        Bw_qty = bw_value/1000
                    cust_uom = cust_value
                elif (cust_unit.lower() == 'kg' and bw_unit.lower() == 'g') or (cust_unit.lower() == 'cm' and bw_unit.lower() == 'm') or (cust_unit.lower()  == "ml" and bw_unit.lower() in (["l","ltr"])):
                    if (cust_unit.lower() == 'cm' and bw_unit.lower() == 'm'):
                        cust_uom = cust_value/100
                    else:
                        cust_uom = cust_value/1000
                    Bw_qty = bw_value
            else: ###liter
                liter_search = re.search(r'(?i)(\d+)(l|LITRE|LT)', row["Part Description"])
                liter_unit = 'l'
                liter_value = liter_search.group(1) if liter_search else 1
                if (cust_unit.lower() == liter_unit.lower() or (len(cust_unit) >= len(liter_unit) and cust_unit[0:len(liter_unit)].lower() == liter_unit.lower()) or (len(cust_unit) < len(liter_unit) and liter_unit[0:len(cust_unit)].lower() == cust_unit.lower())):
                    Bw_qty = liter_value
                    cust_uom = cust_value
        df2.at[index, 'Cust_Size_Bag Comments'] = 'Converted by customer UOM'
    elif re.match(r'^([a-zA-Z]+)([0-9]+)$',row['Cust UOM']):
        ##get the number
        cust_uom = re.search(r'^([a-zA-Z]+)([0-9]+)$', row['Cust UOM']).group(2)
        df2.at[index, 'Cust_Size_Bag Comments'] = 'Converted by customer UOM'
    elif re.match(r'^([0-9]+)([a-zA-Z]+)$',row['Cust UOM']):
        cust_uom = re.search(r'^([0-9]+)([a-zA-Z]+)$', row['Cust UOM']).group(1)
        df2.at[index, 'Cust_Size_Bag Comments'] = 'Converted by customer UOM'
    elif re.search(r'(?i)(\d+)\s*X\s*(\d+)\s*([a-zA-Z]+)' , row['Cust UOM']):
        #6 X 1 Ltr
        cust_uom = re.search(r'(?i)(\d+)\s*X\s*(\d+)\s*([a-zA-Z]+)', row['Cust UOM']).group(1)
        df2.at[index, 'Cust_Size_Bag Comments'] = 'Converted by customer UOM'
        #update Bw_qty
        bracket = re.search(r'\((.*)\)' , row["Part Description"])
        Bw_qty = bracket if (bracket and bracket.group(1).isdigit()) else 1
        df2.at[index, 'BW_Size_Bag Comments'] = 'Update BW by matching customer uom'
    elif re.search(r'(?i)(\d+)\s*(\w+)+X\s*(\d+)', row['Cust UOM']):
        #20PX50     
        cust_uom = re.search(r'(?i)(\d+)\s*(\w+)+X\s*(\d+)', row['Cust UOM']).group(1)
        df2.at[index, 'Cust_Size_Bag Comments'] = 'Converted by customer UOM'
        #update Bw_qty
        bracket = re.search(r'\((.*)\)' , row["Part Description"])
        Bw_qty = bracket if (bracket and bracket.group(1).isdigit()) else 1
        df2.at[index, 'BW_Size_Bag Comments'] = 'Update BW by matching customer uom'
                    
    df2.at[index, 'Cust_Size'] = cust_uom
    df2.at[index, 'BW_Size'] = Bw_qty
   


# In[12]:


###convert for customer
### testing for customer UOM cdescription
import re
for index, row in df2.iterrows():
    cust_uom = 1
    Bw_qty = 1
    if not (pd.isna(row['Cust Description'])): 
        if re.search(r'\((.*)\)' , row['Cust Description']) :
            cust_qty = re.search(r'\((.*)\)' , row['Cust Description'])
            ## 108 X 108 X 76
            if (re.search(r'(?i)\d+\s*X\s*\d+\s*X\s*\d+',cust_qty.group(1))):
                continue
            elif (re.search(r'(\d+)\s*\.\s*(\d+)X(\w+)',cust_qty.group(1))):
                num1 = re.match(r'(\d+)\s*\.\s*(\d+)X(\w+)',cust_qty.group(1)).group(1)
                right = re.match(r'(\d+)\s*\.\s*(\d+)X(\w+)',cust_qty.group(1)).group(3)
                num2 = re.match(r'^(\d+)',right).group(1)
                if (int(num1) <= 0):
                    num1 = 1        
                if (int(num2) <= 0):
                    num2 = 1
                cust_uom = int(num1)*int(num2)
                if cust_uom > 1000:
                    cust_uom = 1
            elif re.match(r'(\d+)X(\w+)',cust_qty.group(1)):
                num1 = re.match(r'(\d+)X(\w+)',cust_qty.group(1)).group(1)
                right = re.match(r'(\d+)X(\w+)',cust_qty.group(1)).group(2)
                num2 = re.match(r'^(\d+)',right).group(1) 
                if (int(num1) <= 0):
                    num1 = 1        
                if (int(num2) <= 0):
                    num2 = 1
                cust_uom = int(num1)*int(num2)  
                if cust_uom > 1000:
                    cust_uom = 1                   
            elif (re.search(r'-',cust_qty.group(1))):
                cust_uom = 1               
            elif (re.search(r'^\s*(\d+)\s*$',cust_qty.group(1))):
                cust_uom = re.search(r'^\s*(\d+)\s*$',cust_qty.group(1)).group() 
            elif (re.search(r'(\w+)(\s*)=(\s*)(\d+)',cust_qty.group(1))):
                #if BAG K/TIDY OSO WHITE 9UM 27L (ROLL=50) 50
                cust_uom =  re.search(r'(\w+)(\s*)=(\s*)(\d+)',cust_qty.group(1)).group(4)
                
            ###of group
            elif (re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',cust_qty.group(1))):
                of_qty = re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',cust_qty.group(1))
                left = of_qty.group(1)
                right = of_qty.group(3)
                cust_uom = int(left)*int(right) 
            elif (re.search(r'(?i)(\d+)\s*(of)\s+',cust_qty.group(1))):
                cust_uom = re.search(r'(?i)(\d+)\s*(of)\s+',cust_qty.group(1)).group(1)
            elif (re.search(r'(?i)(\d+)\s*(of)$',cust_qty.group(1))):
                cust_uom = re.search(r'(?i)(\d+)\s*(of)$',cust_qty.group(1)).group(1)
            elif (re.search(r"(?i)\s+of(\s*)(\d+)",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)\s+of(\s*)(\d+)",cust_qty.group(1)).group(2) 
            elif (re.search(r"(?i)^of(\s*)(\d+)",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)^of(\s*)(\d+)",cust_qty.group(1)).group(2)
            ##3 pk$
            elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",cust_qty.group(1)).group(1)
            ##3 pk #
            elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",cust_qty.group(1)).group(1)
            #pack 100
            elif (re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",cust_qty.group(1)).group(3) 
            ##$pack 100
            elif (re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",cust_qty.group(1)).group(3) 
            elif (re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)",cust_qty.group(1)).group(2)
            elif (re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)",cust_qty.group(1))):
                cust_uom = re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)",cust_qty.group(1)).group(1)
            #elif (re.search(r'(\D*)(\d+)(\D*)$',cust_qty.group(1))):
                #cust_uom = re.search(r'(\D*)(\d+)(\D*)$',cust_qty.group(1)).group(2) 
            else:
                df2.at[index, 'Comments'] = 'Default cust uom at 1'
        ## if there is only one bracket
        elif re.search(r'\((.*)' , row['Cust Description']) :           
        ## eg: PAPER TOILET UNI TORK RL 850SHT 1ST4 (48
            inner = re.search(r'\((.*)' , row['Cust Description']).group(1)
            if inner.isdigit():
                cust_uom = inner     
        elif (re.search(r'\d+\s*X\s*\d+\s*X\s*\d+',row['Cust Description'])):
            continue
        elif (re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)", row['Cust Description'])): 
            ##eg:bag/10
            cust_uom = re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)", row['Cust Description']).group(2)
        elif (re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)", row['Cust Description'])): 
            #2/pk
            cust_uom = re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)", row['Cust Description']).group(1)        
        ###no bracket
        
        ####of group but of is not inside the word   
        elif (re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',row['Cust Description'])):
            of_qty = re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',row['Cust Description'])
            left = of_qty.group(1)
            right = of_qty.group(3)
            cust_uom = int(left)*int(right) 
        elif (re.search(r'(?i)(\d+)\s*(of)\s+',row['Cust Description'])):
            cust_uom = re.search(r'(?i)(\d+)\s*(of)\s+',row['Cust Description']).group(1)
        elif (re.search(r'(?i)(\d+)\s*(of)$',row['Cust Description'])):
            cust_uom = re.search(r'(?i)(\d+)\s*(of)$',row['Cust Description']).group(1)
        elif (re.search(r"(?i)\s+of(\s*)(\d+)",row['Cust Description'])):
            cust_uom = re.search(r"(?i)\s+of(\s*)(\d+)",row['Cust Description']).group(2) 
        elif (re.search(r"(?i)^of(\s*)(\d+)",row['Cust Description'])):
            cust_uom = re.search(r"(?i)^of(\s*)(\d+)",row['Cust Description']).group(2)
        ##### of group end
        elif (re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Cust Description'])): 
            # Pack 100
            cust_uom = re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Cust Description']).group(3)
        elif (re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Cust Description'])):
            cust_uom = re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Cust Description']).group(3)
        
        elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",row["Cust Description"])):
            ##3PK #
            cust_uom = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",row["Cust Description"]).group(1)           
        elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",row["Cust Description"])):
        ##3PK$#
            cust_uom = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",row["Cust Description"]).group(1)
        
        elif (re.search(r'(\w+)(\s*)=(\s*)(\d+)',row['Cust Description'])):
            cust_uom = re.search(r'(\w+)(\s*)=(\s*)(\d+)',row['Cust Description']).group(4)
        else:
            if row['Cust_Size'] == 0:
                cust_uom = 1 
                df2.at[index, 'Comments'] = 'Default cust uom at 1'
    else:
        if row['Cust_Size'] == 0:
            cust_uom = 1
            df2.at[index, 'Comments'] = 'Default cust uom at 1'
    
    ##update with the customer and BW size 
    if df2.at[index, 'Cust_Size'] == 0:
        df2.at[index, 'Cust_Size'] = cust_uom
    else:
        if int(row['Cust_Size']) != int(cust_uom):
            #print(row['Cust Description'],row['Cust_Size']," custo uom: ", cust_uom)
            if(int(row['Cust_Size']) != 1 and cust_uom != 1):
                df2.at[index, 'Cust_Size'] = min(int(row['Cust_Size']),int(cust_uom))
            elif(int(row['Cust_Size']) == 1 and cust_uom != 1):
                if ((int(cust_uom) / int(row['Cust_Size'])) > 5000 or (int(cust_uom) / int(row['Cust_Size'])) < 0.0002):
                    df2.at[index, 'Cust_Size'] = 1
                else:
                    df2.at[index, 'Cust_Size'] = int(cust_uom)
            elif(int(row['Cust_Size']) != 1 and cust_uom == 1):
                if ((int(cust_uom) / int(row['Cust_Size'])) > 5000 or (int(cust_uom) / int(row['Cust_Size'])) < 0.0002):
                    df2.at[index, 'Cust_Size'] = 1
                else:
                    df2.at[index, 'Cust_Size'] = int(row['Cust_Size'])
            else:
                df2.at[index, 'Cust_Size'] = cust_uom
        else:
            df2.at[index, 'Cust_Size'] = cust_uom


# In[13]:


import requests
import urllib
import requests_html
from requests_html import HTML
from requests_html import HTMLSession
from bs4 import BeautifulSoup
def get_results(query): 
    query = urllib.parse.quote_plus(query)
    response = get_source("https://www.google.com/search?q=" + query) 
    return response
def parse_results(response):    
    css_identifier_result = ".tF2Cxc"
    css_identifier_title = "h3"
    css_identifier_link = ".yuRUbf a"
    css_identifier_text = ".IsZvec"    
    results = response.html.find(css_identifier_result)
    output = []    
    for result in results:
        item = result.find(css_identifier_link, first=True).attrs['href']
        output.append(item)        
    return output
def get_source(url):
    try:
        session = HTMLSession()
        response = session.get(url,verify = False)
        return response
    except requests.exceptions.RequestException as e:
        print(e)
def google_search(query):
    response = get_results(query)
    return parse_results(response)


# In[14]:


##testing for BW UOM
for index, row in df2.iterrows():
    row['Part Description'] = str(row['Part Description'])
    Bw_qty = 1
    if (str(row['Part Number']).upper() in (["UTQ","NLA","MIR","POA"])):
        Bw_qty = 0
        df2.at[index, 'BW_Size_Bag Comments'] = 'Unpriced'
    elif row['Part Number'] == "":
        Bw_qty = 0
        df2.at[index, 'BW_Size_Bag Comments'] = 'No Part Number'
    elif not (pd.isna(row['Part Description'])): 
            if (pd.isna(row['Price UOM'])):
                df2.at[index, 'BW_Size_Bag Comments'] = 'No BW Price UOM'
                continue
            elif row['Price UOM'] in (["100","1000"]):
                Bw_qty = int(row['Price UOM'])  
            else:
                Bw_unit = re.search(r'\((.*)\)' , row['Part Description']) 
                if Bw_unit: 
                    if re.search(r'\((.*)\)' , row['Part Description']).group(1).isdigit():
                        Bw_qty = re.search(r'\((.*)\)' , row['Part Description']).group(1)
                    ## 108 X 108 X 76
                    elif (re.search(r'(?i)\d+\s*X\s*\d+\s*X\s*\d+',Bw_unit.group(1))):
                        continue
                        
                    elif (re.search(r'(\d+)\s*\.\s*(\d+)X(\w+)',Bw_unit.group(1))):
                        Bw_num1 = re.match(r'(\d+)\s*\.\s*(\d+)X(\w+)',Bw_unit.group(1)).group(1)
                        right = re.match(r'(\d+)\s*\.\s*(\d+)X(\w+)',Bw_unit.group(1)).group(3)
                        Bw_num2 = re.match(r'^(\d+)',right).group(1)
                        if (int(Bw_num1) <= 0):
                            Bw_num1 = 1        
                        if (int(Bw_num2) <= 0):
                               Bw_num2 = 1
                            
                        if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                            if (round(int(row['Cust_Size'])) == round(int(Bw_num1))):
                                Bw_qty = int(Bw_num1) * int(Bw_num2)
                            elif (round(int(row['Cust_Size'])) == round(int(Bw_num2))):
                                Bw_qty = int(Bw_num1) * int(Bw_num2)
                            elif (round(int(row['Cust_Size'])) == round(int(Bw_num1) * int(Bw_num2))):
                                Bw_qty = int(Bw_num1) * int(Bw_num2)
                            elif (((int(Bw_num1) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(Bw_num1) == 0))) and (int(Bw_num1) != 0):
                                Bw_qty = Bw_num1
                            elif (((int(Bw_num2) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(Bw_num2) == 0))) and (int(Bw_num2) != 0):
                                Bw_qty = Bw_num2
                            elif row['Part Description'].endswith(str(row['Cust_Size'])):
                                Bw_qty = int(row['Cust_Size']) 
                            else:
                                Bw_qty = 1
                        else:
                            Bw_qty = 1    
                           
                    elif re.match(r'(\d+)X(\w+)',Bw_unit.group(1)):
                        Bw_num1 = re.match(r'(\d+)X(\w+)',Bw_unit.group(1)).group(1)
                        right = re.match(r'(\d+)X(\w+)',Bw_unit.group(1)).group(2)
                        Bw_num2 = re.match(r'^(\d+)',right).group(1) 
                        if (int(Bw_num1) <= 0):
                            Bw_num1 = 1        
                        if (int(Bw_num2) <= 0):
                            Bw_num2 = 1
                        if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                            if (round(int(row['Cust_Size'])) == round(int(Bw_num1))):
                                Bw_qty = int(Bw_num1) * int(Bw_num2)
                                
                            elif (round(int(row['Cust_Size'])) == round(int(Bw_num2))):
                                Bw_qty = int(Bw_num1) * int(Bw_num2)
                                
                            elif (round(int(row['Cust_Size'])) == round(int(Bw_num1) * int(Bw_num2))):
                                Bw_qty = int(Bw_num1) * int(Bw_num2)
                                
                            elif (((int(Bw_num1) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(Bw_num1) == 0))) and (int(Bw_num1) != 0):
                                Bw_qty = Bw_num1
                                
                            elif (((int(Bw_num2) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(Bw_num2) == 0))) and (int(Bw_num2) != 0):
                                Bw_qty = Bw_num2
                            elif row['Part Description'].endswith(str(row['Cust_Size'])):
                                Bw_qty = int(row['Cust_Size'])  
                            else:
                                Bw_qty = 1
                        else:
                            Bw_qty = 1          
                        
                    
                  
                    elif (re.search(r'-',Bw_unit.group(1))):
                        Bw_qty = 1               
                    elif (re.search(r'^\s*(\d+)\s*$',Bw_unit.group(1))):
                        if int(re.search(r'^\s*(\d+)\s*$',Bw_unit.group(1)).group()) < 1000:
                            Bw_qty = re.search(r'^\s*(\d+)\s*$',Bw_unit.group(1)).group() 
                    elif (re.search(r'(\w+)(\s*)=(\s*)(\d+)',Bw_unit.group(1))):
                        #72MM=8X9MMSEG
                        if re.search(r'(\w+)(\s*)=(\s*)(\d+)X(\d+)',Bw_unit.group(1)):
                            Bw_front = re.search(r'(\w+)(\s*)=(\s*)(\d+)X(\d+)',Bw_unit.group(1)).group(4)
                            Bw_back = re.search(r'(\w+)(\s*)=(\s*)(\d+)X(\d+)',Bw_unit.group(1)).group(5)
                            if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                                if (round(int(row['Cust_Size'])) == round(int(Bw_front))) or (round(int(row['Cust_Size'])) == round(int(Bw_back))) or (round(int(row['Cust_Size'])) == round(int(Bw_front) * int(Bw_back))):
                                    Bw_qty = int(Bw_front) * int(Bw_back)
                                else:
                                    Bw_qty = 1
                        else:
                        #if BAG K/TIDY OSO WHITE 9UM 27L (ROLL=50) 50
                            Bw_qty =  re.search(r'(\w+)(\s*)=(\s*)(\d+)',Bw_unit.group(1)).group(4)    
                    ###of group
                    elif (re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',Bw_unit.group(1))):
                        of_qty = re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',Bw_unit.group(1))
                        left = of_qty.group(1)
                        right = of_qty.group(3)
                        Bw_qty = int(left)*int(right) 
                    elif (re.search(r'(?i)(\d+)\s*(of)\s+',Bw_unit.group(1))):
                        Bw_qty = re.search(r'(?i)(\d+)\s*(of)\s+',Bw_unit.group(1)).group(1)
                    elif (re.search(r'(?i)(\d+)\s*(of)$',Bw_unit.group(1))):
                        Bw_qty = re.search(r'(?i)(\d+)\s*(of)$',Bw_unit.group(1)).group(1)
                    elif (re.search(r"(?i)\s+of(\s*)(\d+)",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)\s+of(\s*)(\d+)",Bw_unit.group(1)).group(2) 
                    elif (re.search(r"(?i)^of(\s*)(\d+)",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)^of(\s*)(\d+)",Bw_unit.group(1)).group(2)
                        
                    ##3 pk$
                    elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",Bw_unit.group(1)).group(1)  
                    ##3 pk #
                    elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",Bw_unit.group(1)).group(1)
                    ## 3 pkX20
                    elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|pks)\s*X\d+",Bw_unit.group(1))):
                        search_front = Bw_qty = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|pks)\s*X\d+",Bw_unit.group(1)).group(1)
                        search_back = Bw_qty = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|pks)\s*X\d+",Bw_unit.group(1)).group(4)      
                        if int(row['Cust_Size']) == int(search_front):
                            Bw_qty = int(search_front) * int(search_back)
                        elif int(row['Cust_Size']) == int(search_back):
                            Bw_qty = int(search_front) * int(search_back)
                        else:
                            Bw_qty = int(search_front)
                   
                    ##pack 100
                    elif (re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",Bw_unit.group(1)).group(3) 
                    ##^pack 100
                    elif (re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)",Bw_unit.group(1)).group(3) 
                    elif (re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)",Bw_unit.group(1)).group(2)

                    elif (re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)",Bw_unit.group(1))):
                        Bw_qty = re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)",Bw_unit.group(1)).group(1)

                    #elif (re.search(r'(\D*)(\d+)(\D*)$',Bw_unit.group(1))):
                        #Bw_qty = re.search(r'(\D*)(\d+)(\D*)$',Bw_unit.group(1)).group(2)  

                    
                    
                    
                ## if there is only one bracket
                elif re.search(r'\((.*)' , row['Part Description']) :           
                ## eg: PAPER TOILET UNI TORK RL 850SHT 1ST4 (48
                    inner = re.search(r'\((.*)' , row['Part Description']).group(1)
                    if inner.isdigit():
                        Bw_qty = inner 
                ###no bracket
                elif (re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)", row['Part Description'])): 
                    ##eg:bag/10
                    Bw_qty = re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)", row['Part Description']).group(2)
                    
                elif (re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)", row['Part Description'])): 
                    #2/pk
                    Bw_qty = re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)", row['Part Description']).group(1)
                
                ####of group but of is not inside the word   
                elif (re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',row['Part Description'])):
                    of_qty = re.search(r'(?i)(\d+)\s*(Of)\s*(\d+)',row['Part Description'])
                    left = of_qty.group(1)
                    right = of_qty.group(3)
                    Bw_qty = int(left)*int(right) 
                elif (re.search(r'(?i)(\d+)\s*(of)\s+',row['Part Description'])):
                    Bw_qty = re.search(r'(?i)(\d+)\s*(of)\s+',row['Part Description']).group(1)
                elif (re.search(r'(?i)(\d+)\s*(of)$',row['Part Description'])):
                    Bw_qty = re.search(r'(?i)(\d+)\s*(of)$',row['Part Description']).group(1)
                elif (re.search(r"(?i)\s+of(\s*)(\d+)",row['Part Description'])):
                    Bw_qty = re.search(r"(?i)\s+of(\s*)(\d+)",row['Part Description']).group(2) 
                elif (re.search(r"(?i)^of(\s*)(\d+)",row['Part Description'])):
                    Bw_qty = re.search(r"(?i)^of(\s*)(\d+)",row['Part Description']).group(2)
                ##### of group end
                
                elif (re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Part Description'])): 
                    #Pack 100
                    Bw_qty = re.search(r"(?i)[\s+|,](pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Part Description']).group(3)
                elif (re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Part Description'])): 
                    #Pack 100
                    Bw_qty = re.search(r"(?i)^(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)(\s*)(\d+)", row['Part Description']).group(3)
                elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",row["Part Description"])):
                    ##3PK$
                    Bw_qty = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)$",row["Part Description"]).group(1)
                    
                elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",row["Part Description"])):
                    ##3PK #
                    Bw_qty = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\s+",row["Part Description"]).group(1)
                #3Pk X20
                elif (re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|pks)\s*X(\d+)",row["Part Description"])):
                    search_front = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|pks)\s*X(\d+)",row["Part Description"]).group(1)
                    search_back = re.search(r"(?i)(\d+)(\s*)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|pks)\s*X(\d+)",row["Part Description"]).group(4)       
                    if int(row['Cust_Size']) == int(search_front):
                        Bw_qty = int(search_front) * int(search_back)
                    elif int(row['Cust_Size']) == int(search_back):
                        Bw_qty = int(search_front) * int(search_back)
                    else:
                        Bw_qty = int(search_front)
                
                elif (re.search(r'(\w+)(\s*)=(\s*)(\d+)',row['Part Description'])):
                    Bw_qty = re.search(r'(\w+)(\s*)=(\s*)(\d+)',row['Part Description']).group(4)
                elif (re.search(r'(\d+)-(\d+)',row['Part Description'])):
                    Bw_qty = 1  
                elif  (re.search(r'\d+\s*X\s*\d+\s*X\s*\d+', row['Part Description'])):
                    continue    
                elif (re.search(r"(?i)(\d+\.\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe)$", row['Part Description'])):
                    value = re.search(r"(?i)(\d+\.\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe)$", row['Part Description']).group(1)
                    value = int(float(value))
                    if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                        if (round(int(row['Cust_Size'])) == round(float(value))):
                            Bw_qty = value
                        elif (((int(value) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(value) == 0))) and (int(value) != 0):
                            Bw_qty = value
                        else:
                            Bw_qty = 1
                    else:
                        Bw_qty = 1
                elif (re.search(r"(?i)((\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))$", row['Part Description'])):
                    measurement = re.search(r"(?i)((\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))$", row['Part Description']).group(1)
                    ##get the value
                    value = int(re.search(r'\d+', measurement).group(0))
                    if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                        if measurement in row['Cust Description']:
                            if value == row['Cust_Size']:
                                Bw_qty = value
                            else:
                                Bw_qty = 1
                        elif (round(int(row['Cust_Size'])) == round(int(value))):
                            Bw_qty = value
                        elif (((int(value) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(value) == 0))) and (int(value) != 0):
                            Bw_qty = value
                        else:
                            Bw_qty = 1
                    else:
                        Bw_qty = 1
                elif (re.search(r"(?i)((\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))", row['Part Description'])):
                    measurement = re.search(r"(?i)((\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))", row['Part Description']).group(1)
                    ##get the value
                    value = int(re.search(r'\d+', measurement).group(0))
                    if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                        try:
                            if re.search(r"(?i)(\d+\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))", row['Cust Description']).group(1) in row['Part Description']:
                                cust_measurement = re.search(r"(?i)(\d+\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))", row['Cust Description']).group(1)
                                Bw_qty = row['Cust_Size']
                                
                        except:
                            if measurement in row['Cust Description']:
                                if value == row['Cust_Size']:
                                    Bw_qty = value
                                else:
                                    Bw_qty = 1
                                
                            elif (round(int(row['Cust_Size'])) == round(int(value))):
                                Bw_qty = value
                            elif (((int(value) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(value) == 0))) and (int(value) != 0):
                                Bw_qty = value
                            else:
                                Bw_qty = 1
                    else:
                        Bw_qty = 1               
    else: #no description
        if row['Price UOM'] in (["100","1000"]):
            Bw_qty = int(row['Price UOM'])
        else:
            Bw_qty = 1
            df2.at[index, 'Comments'] = 'No BW Description'
            
            
    ###final check        
    ## Bw_qty != cust_uom && Bw_qty <> 1
    if (int(Bw_qty) != int(row['Cust_Size']) and Bw_qty != 1 and Bw_qty != 0 ):
         #Bw uom = each but qty != 1
        if row['Price UOM'] in (["EA","Each","ea","EACH"]) or row['Cust UOM'] in (["EA","Each","ea","EACH"]):
            if(row['Cust_Size'] == 1):
                try:
                    if (Bw_qty in row['Cust UOM'] or row['Cust Description'].endswith(Bw_qty)):
                        df2.at[index, 'Cust_Size'] = Bw_qty 
                        df2.at[index, 'Cust_Size_Bag Comments'] = 'Map customer size bag with BW quantity'
                    elif (row['Cust_Size'] not in (["EA","Each","ea","EACH"])):
                        df2.at[index, 'Comments'] = 'Check Customer UOM for more information'
                except:
                    pass
        #if both are not each
        if( (row['Cust UOM'] not in (["EA","Each","ea","EACH"])) and(row['Price UOM'] not in (["EA","Each","ea","EACH"])) ):
            if(row['Cust_Size'] == 1):
                #end with the bw qty
                try:
                    if (Bw_qty in row['Cust UOM'] or row['Cust Description'].endswith(Bw_qty)):
                        df2.at[index, 'Cust_Size'] = Bw_qty 
                        df2.at[index, 'Cust_Size_Bag Comments'] = 'Map customer size bag with BW quantity'
                    else:
                        df2.at[index, 'Comments'] = 'Check Customer UOM for more information'
                except:
                    pass
    ## Bw_qty != cust_uom && Bw_qty == 1: web-scraping
    elif (int(Bw_qty) != int(row['Cust_Size']) and Bw_qty == 1):
        if (row['Cust UOM']) not in  (["EA","Each","ea","EACH"]):
            if (row['Price UOM']) not in  (["EA","Each","ea","EACH","PAIR"]):
                #if ((row['Price UOM']) == (row['Cust UOM'])):
                if str(row['Part Number']).lower() == "buyin":
                    df2.at[index, 'Comments'] = 'Check BUYIN UOM for more information'
                elif str(row['Part Number']).isdigit():
                    ####web scraping
                    row['Part Number'] = str(row['Part Number']).zfill(8)
                    try:
                        url = google_search('Blackwoods' + row['Part Number'])[0]
                        print(url)
                    except:
                        Bw_qty = row['Cust_Size']
                        df2.at[index, 'Comments'] = "Double check BW UOM as it is the invalid part number"
                        continue
                        
                    session = requests.Session()
                    session.headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.1.2222.33 Safari/537.36",
                    "Accept-Encoding": "*",
                    "Connection": "keep-alive"
                    }
                    ##iferror: pass
                    try:
                        r = session.get(url)
                        soup = BeautifulSoup(r.text, "html.parser")
                    except:
                        continue
                    
                    if soup.find_all('h1', class_='productCaption'):
                        content = soup.find_all('h1', class_='productCaption')[0].text
                        #bracket
                        if re.search(r'\((.*)\)' , content) :        
                            inner = re.search(r'\((.*)\)' , content).group(1)
                            if inner.isdigit():
                                Bw_qty = inner 
                            elif re.search(r'(?i)of\s+(\d+)' , inner):
                                Bw_qty = re.search(r'(?i)of\s+(\d+)',inner).group(1)
                            elif re.search(r'(?i)(\d+)\s+of' , inner):
                                Bw_qty = re.search(r'(?i)(\d+)\s+of' , inner).group(1)
                        #single bracket
                        elif re.search(r'\((.*)',content) :        
                            inner = re.search(r'\((.*)', content).group(1)
                            if inner.isdigit():
                                Bw_qty = inner 
                        #bag of 10
                        elif (re.search(r"(?i)\s+of(\s*)(\d+)",content)):
                            if not (re.search(r"(?i)\s+of(\s*)(\d+)\s*Wiper",content)) and not (re.search(r"(?i)\s+of(\s*)(\d+)\s*sheet",content)) :
                                Bw_qty = (re.search(r"(?i)\s+of(\s*)(\d+)",content)).group(2) 
                                
                        #10 of bag
                        elif (re.search(r"(?i)(\d+)\s*of",content)):
                            Bw_qty = re.search(r"(?i)(\d+)\s*of",content).group(1) 
                        #bag 500
                        elif (re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|ctn)(\s*)(\d+)", content)):
                             Bw_qty = re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|ctn)(\s*)(\d+)", content).group(3)
                        #100 pack
                        elif (re.search(r"(?i)(\d+)\s*(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)", content)):
                             Bw_qty = re.search(r"(?i)(\d+)\s*(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)", content).group(1)
                        #bag/3
                        elif (re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)",content)):
                             Bw_qty =re.search(r"(?i)(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn)\/(\d+)",content).group(2)
                        #1/bag
                        elif (re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)",content)):
                            Bw_qty = re.search(r"(?i)(\d+)\D*\/(pkt|pack|pk|box|bx|bag|bg|set|carton|pads|pe|ctn|roll)",content).group(1)
                        #Post-it® Easel Pads White 635x774mm 559 
                        elif (re.search(r'(?i)(\d+)\s*X\s*(\d+)',content)):
                            if row['Cust_Size'] == re.search(r'(?i)(\d+)\s*X\s*(\d+)',content).group(1) or row['Cust_Size'] == re.search(r'(?i)(\d+)\s*X\s*(\d+)',content).group(2):
                                Bw_qty = row['Cust_Size']
                            else:
                                df2.at[index, 'Comments'] = 'Default BW = cust unit, checking BW website for more information'
                                Bw_qty = row['Cust_Size']
                        ## 135x165x240
                        elif (re.search(r'(?i)\d+\s*X\s*\d+\s*X\s*\d+',content)):
                            df2.at[index, 'Comments'] = 'Default BW = cust unit, checking BW website for more information'
                            Bw_qty = row['Cust_Size']
                        elif (re.search(r"(?i)((\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))", content)):
                            measurement = re.search(r"(?i)((\d+)\s*(m|cm|g|kg|ml|l|LITRE|LT|Ltr|Lt|ltr|LTRS|GM|wipe))", content).group(1)
                            value = int(re.search(r'\d+', measurement).group(0))
                            if (row['Cust_Size'] != 1 and row['Cust_Size'] != 0):
                                if measurement in row['Cust Description']:
                                    if value == row['Cust_Size']:
                                        Bw_qty = value
                                    else:
                                        Bw_qty = 1
                                elif (round(int(row['Cust_Size'])) == round(int(value))):
                                    Bw_qty = value
                                elif (((int(value) % int(row['Cust_Size']) == 0) or (int(row['Cust_Size']) % int(value) == 0))) and (int(value) != 0):
                                    Bw_qty = value
                                else:
                                    Bw_qty = 1
                            else:
                                Bw_qty = 1
                    print("Cust: ",row['Cust Description'],"Size: ",row['Cust_Size'], "Sort Key: ",row['Sort Key'], "BW: ",Bw_qty)
                    df2.at[index, 'BW_Size_Bag Comments']='web scraping BW size bag'
    #if Bw_qty == 0:
        #Bw_qty = 1
    df2.at[index, 'BW_Size'] = Bw_qty   


# In[15]:


##Conversion
for index, row in df2.iterrows():
    if float(row['BW_Size']) == 0:
        df2.loc[index,'conversion'] = 0
    else:
        conv = float(row['Cust_Size'])/float(row['BW_Size'])
        df2.loc[index,'conversion'] = float(conv)


# In[16]:


df2.to_excel("conversion_v1.xlsx",sheet_name='uom_conversion',index=False)  


# In[ ]:




