import pandas as pd
import numpy as np
import os
import win32com.client as win32




iari_d = pd.read_csv(r"V:\COMPARTILHADO_CSC-SSE_NSIF\NP-2\GPI\4 - Apoio Administrativo\4.4 - Monte "
                         "Albuquerque\Desenv\Pedro\Relatórios Mensais\iari.txt", sep="|", skiprows=91, encoding="ISO-8859-1", usecols=lambda x: x not in ['Unnamed: 0', 'Unnamed: 21'], skip_blank_lines=True)
iari_d.rename(columns=lambda x: x.strip(), inplace=True)
iari_d.columns = iari_d.columns.str.strip()
iari_d.dropna(axis=0, inplace=True)
iari_d = iari_d.cont[iari_d['Ordem'] != "Ordem"]

iari_d['tipo_nota'] = np.where(iari_d['CóMd'] == "D", pd.DateOffset(days=540),
                                  np.where(iari_d['CóMd'] == "C", pd.DateOffset(days=360),
                                           np.where(iari_d['CóMd'] == "B", pd.DateOffset(days=120),
                                                    np.where(iari_d['CóMd'] == "A",
                                                             pd.DateOffset(days=15),
                                                             False))))
print(iari_d['Ordem'][88])