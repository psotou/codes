import pandas as pd
import numpy as np 
import os, glob

data = input('Nombre del archivo (incluir extensión):')
mes = input('Mes de interés (tal cual está escrito en el excel):')
col = input('Columnas de interés (escribir tal cual aparecen):').split(', ')

df = pd.read_excel(( os.getcwd() + '/' + data ).replace('\\','/'), header=0 )
df = df[df[mes]==1] #hasta mayo
df = df[col] #['Cartera', 'Área Funcional', 'Nivel', 'Denominación nivel']
df['Nivel'] = df.Nivel.str.slice(start=4)
df0 = df.copy()

#gerencias: ajustar agrupación de estas. 
proy = ['andina', 'chuqui subte', 'talabre', 'rajo inca', 'nuevo nivel mina', 'carén|caren', 'óxidos|súlfuros', 'relaves', 'vicepresidencia', 'recursos humanos', 'administración', 'riesgo', 'construc', 'sustentabi', 'seguridad', 'ácido|acido', 'agua desalada']

#--------------------------------------------------------------------------------------------------------
l = list(set(df[col[1]]))
l.sort()
it = ['Máximo nivel', 'Dotación máximo nivel', 'Dotación']

o, zz = [], []
for i in range(len(proy)):    
    for j in range(len(l)):        
        w = df[df[col[0]].str.contains(proy[i], case=False)]
        x = w[w[col[1]].str.contains(l[j], case=False)]

        if (x.Nivel == 'E').any():
            xmin = 'E'
        else:
            xmin = x.Nivel.min()#máximo nivel

        y = x[x.Nivel == xmin].Nivel.count()#dotación máximo nivel
        dot = len(x)#dotación

        f = np.array([xmin, y, dot])

        g = pd.DataFrame(index=it)
        g[l[j]] = f
        o.append(g)
        oo = pd.concat(o, axis=1)
        zz.append(oo)

d = zz[len(zz)-1]

aa = []
for i in range(len(proy)):
    a = d.iloc[:, i*len(l):(i+1)*len(l)]
    aa.append(a)
    ax = pd.concat(aa, axis=0)

ax.reset_index(level=0, inplace=True)

b = [None]*len(ax)

for i in range(len(proy)):
    b[i*3] = proy[i]

axx = ax.set_index(pd.Index(b))

axx.to_excel('dotación_nuevo.xlsx')
