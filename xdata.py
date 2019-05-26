import pandas as pd
import numpy as np 
import os, glob

print('NOTA: el nombre de las columnas con las que se cruzan los datos son "Proyecto", "Área Funcional", "Nivel" y "Denominación nivel". Si dichas columnas tienen algún cambio (por ejemplo, no tienen tilde o tienen mayúsculas en vez de minúsculas) cambie los nombres para que queden tal como los mencionados anteriormente.')

# data = 'Libro3.xlsx'
data = input('Nombre del archivo (incluir extensión):')
mes = input('Ingrese el mes de interés:')

df = pd.read_excel(( os.getcwd() + '/' + data ).replace('\\','/'), header=0 )
df = df[df[mes]==1] #hasta mayo
df = df[['Proyecto', 'Área Funcional', 'Nivel', 'Denominación nivel']]
#df = df.iloc[:, [0,3,5,6,7]] #selecciona las columnas de interés en caso de que los nombres cambien
df['Nivel'] = df.Nivel.str.slice(start=4)
df0 = df.copy()

#ger = list(set(df.Proyecto)) #Sirve para generar una lista con las gerencias

#gerencias: ajustar agrupación de estas. 
proy = ['andina', 'chuqui', 'talabre', 'rajo inca', 'nuevo nivel mina', 'carén|caren', 'óxidos|súlfuros', 'relaves', 'vicepresidencia', 'recursos humanos', 'administración', 'riesgo', 'construc', 'sustentabi', 'seguridad', 'ácido|acido', 'agua desalada']

#-------------------------------------------------ESTE ES EL FINO-----------------------------------
l = list(set(df0['Área Funcional']))
l.sort()
it = ['Máximo nivel', 'Dotación máximo nivel', 'Dotación']

o, zz = [], []
for i in range(len(proy)):    
    for j in range(len(l)):        
        w = df[df.Proyecto.str.contains(proy[i], case=False)]
        x = w[w['Área Funcional'].str.contains(l[j], case=False)]

        if x.Nivel.isin(['E']).any():
            xmin = 'E'
        else:
            xmin = x.Nivel.min()

        #xmin = x.Nivel.min()#máximo nivel
        y = x[x.Nivel == xmin].Nivel.count()#dotación máximo nivel
        dot = len(x)#dotación

        f = np.array([xmin, y, dot])

        g = pd.DataFrame(index=it)
        g[l[j]] = f
        o.append(g)
        oo = pd.concat(o, axis=1)
        zz.append(oo)

d = zz[len(zz)-1]

with pd.ExcelWriter('dotación por cartera.xlsx') as writer:
    for i in range(len(proy)):
        d.iloc[:, i*20:(i+1)*20].to_excel(writer, sheet_name=proy[i])
