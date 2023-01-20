import xml.etree.ElementTree as ETree
import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

#path = askdirectory(title='Selecione a pasta')

def xml_planilhar(path,nome_plan):
    data = {'DATA HORA EMISSÃO':[],
            'NUMERO NOTA':[],
            'FORNECEDOR': [],
            'CFOP': [],
            'VALOR':[],
            'ID':[],
            }

    os.chdir(path)
    for  file in os.listdir():
        if file.endswith('xml'):


            Tree = ETree.parse(file)
            root = Tree.getroot()
            tag = root.tag
            att = root.attrib
            conta = 0
            __data = {'DATA HORA EMISSÃO':[],
                    'NUMERO NOTA':[],
                    'FORNECEDOR': [],
                    'CFOP': [],
                    'VALOR':[],
                    'ID':[],
            }


            for child in root:
                #print(child.tag,'--',child.attrib, '--',child.text,'CHILD')
                for subchild in child:
                    #print(subchild.tag,'--',subchild.attrib,'--',subchild.text, 'SUBCHILD')
                    if subchild.tag == '{http://www.portalfiscal.inf.br/nfe}infNFe':
                        __data['ID'] += [subchild.attrib['Id'][3:]]
                    for subchild2 in subchild:
                       #print(subchild2.tag,'--',subchild2.attrib,'--',subchild2.text,'SUB2')
                        if subchild2.tag == '{http://www.portalfiscal.inf.br/nfe}ide':
                            for ele in subchild2:
                                if ele.tag == '{http://www.portalfiscal.inf.br/nfe}dhEmi':
                                    __data['DATA HORA EMISSÃO'] += [ele.text]
                                if ele.tag == '{http://www.portalfiscal.inf.br/nfe}natOp':
                                    __data['CFOP'] += [ele.text]
                                if ele.tag == '{http://www.portalfiscal.inf.br/nfe}nNF':
                                    __data['NUMERO NOTA'] += [ele.text]
                        elif subchild2.tag == '{http://www.portalfiscal.inf.br/nfe}emit':
                            for ele in subchild2:
                                if ele.tag == '{http://www.portalfiscal.inf.br/nfe}xNome':
                                    __data['FORNECEDOR'] += [ele.text]
                        for subchild3 in subchild2:
                            if subchild3.tag == '{http://www.portalfiscal.inf.br/nfe}prod':
                                for ele in subchild3:
                                    if ele.tag == '{http://www.portalfiscal.inf.br/nfe}vProd':
                                        conta += float(ele.text)
      
        __data['VALOR'].append(conta)
        for i in __data['VALOR']:
            data['VALOR'].append(i)

        for i in __data['DATA HORA EMISSÃO']:
            data['DATA HORA EMISSÃO'].append(i)

        for i in __data['FORNECEDOR']:
            data['FORNECEDOR'].append(i)
            
        for i in __data['CFOP']:
            data['CFOP'].append(i)

        for i in __data['ID']:
            data['ID'].append(i)

        for i in __data['NUMERO NOTA']:
            data['NUMERO NOTA'].append(i)

    df = pd.DataFrame(data)

    path = askdirectory(title='Selecione a pasta para salvar')
    writer = pd.ExcelWriter(f'{path}/{nome_plan}.xlsx')
    df.to_excel(writer, sheet_name='Teste', index=False, na_rep='NaN')

    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Teste'].set_column(col_idx, col_idx, column_length)

    writer.save()

