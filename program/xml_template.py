# шаблон для файла xml с првоерками
xml_template_clashtests = '''<?xml version="1.0" encoding="UTF-8"?>  
    
<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd" units="m" filename="Проверки" filepath="">  
    <batchtest name="Проверки" internal_name="Проверки" units="m">    
        <clashtests>      
            {0}
        </clashtests>    
        <selectionsets>    
        </selectionsets>  
    </batchtest>
</exchange>'''

# шаблон одной проверки
xml_template_clashtest = '''<clashtest name="{0}" test_type="hard" status="new" tolerance="{1}" merge_composites="1">  
                <linkage mode="none"/>        
                <left>          
                    <clashselection selfintersect="0" primtypes="1">            
                        <locator>lcop_selection_set_tree/Поисковые наборы по моделям/{2}/{3}</locator>          
                    </clashselection>        
                </left>       
                <right>          
                    <clashselection selfintersect="0" primtypes="1">            
                        <locator>lcop_selection_set_tree/Поисковые наборы по моделям/{4}/{5}</locator>          
                    </clashselection>        
                </right>        
                <rules/>      
            </clashtest>
            '''

# шаблон для файла xml с поисковыми наборами
xml_template_selectionsets = '''<?xml version="1.0" encoding="UTF-8"?>  
    
<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd" units="ft" filename="" filepath="">  
    <selectionsets>
        <viewfolder name="Модели" guid="">
            {0}
        </viewfolder>
        <viewfolder name="Поисковые наборы по моделям" guid="">
            {1}
        </viewfolder>
    </selectionsets>
</exchange>'''

# шаблон для папки
xml_template_selectionset_folder = '''<viewfolder name="{0}" guid="">
                {1}
            </viewfolder>
            '''

# шаблон для поискового набора по элементам
xml_template_selectionset = '''<selectionset name="{0}" guid="">  
                    <findspec mode="all" disjoint="0">        
                        <conditions>     
                            <condition test="contains" flags="10">            
                                <property>              
                                    <name internal="LcOaNodeSourceFile">Файл источника</name>            
                                </property>            
                                <value>              
                                    <data type="wstring">{1}</data>            
                                </value>          
                            </condition>                 
                            <condition test="contains" flags="10">            
                                <category>              
                                    <name internal="LcRevitData_Element">Объект</name>            
                                </category>            
                                <property>              
                                    <name internal="LcRevitPropertyElementCategory">Категория</name>            
                                </property>            
                                <value>              
                                    <data type="wstring">{2}</data>            
                                </value>          
                            </condition>        
                        </conditions>        
                        <locator>/</locator>      
                    </findspec>    
                </selectionset>
                '''

# шаблон для поискового набора по моделям
xml_template_selectionset_models = '''<selectionset name="{0}" guid="">  
                <findspec mode="all" disjoint="0">        
                    <conditions>          
                        <condition test="contains" flags="10">            
                            <property>              
                                <name internal="LcOaNodeSourceFile">Файл источника</name>            
                            </property>            
                            <value>              
                                <data type="wstring">{0}</data>            
                            </value>          
                        </condition>                
                    </conditions>        
                    <locator>/</locator>      
                </findspec>    
            </selectionset>
            '''