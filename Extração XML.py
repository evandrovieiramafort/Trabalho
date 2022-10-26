C1 = t.time() 
cab_1 = ["natOp","dhEmi","xNome","xMun","UF",
         "infCpl","cNF","nItem","vProd","vFrete","vDesc","vNF"]

headers = ["Operação","Data","Nome","Cidade","UF",
           "Loja","Número da NF","Itens comprados","Valor da Venda","Valor do Frete","Desconto","Total"]

cab_2 = ["dhEmi","cProd","xProd","qCom","vUnCom","vFrete","vProd","nItemProd","vDesc"]

listas1 = [[] for item in range(len(cab_1))]
listas2 = [[] for item in range(len(cab_2))]

for item in os.listdir():
    if item.endswith(".xml"):
        xml = open(item)
        nfe = minidom.parse(xml)
        
        op = nfe.getElementsByTagName('natOp')
        data = nfe.getElementsByTagName('dhEmi')
        nome = nfe.getElementsByTagName('xNome')
        cidade = nfe.getElementsByTagName('xMun')
        UF = nfe.getElementsByTagName('UF')
        loja = nfe.getElementsByTagName('infCpl')
        NF = nfe.getElementsByTagName('cNF')
        itens = nfe.getElementsByTagName('det')
        prod = nfe.getElementsByTagName('vProd')
        frete = nfe.getElementsByTagName('vFrete')
        desc = nfe.getElementsByTagName('vDesc')
        total = nfe.getElementsByTagName('vNF')
        
        listas1[0].append(op[0].firstChild.data)
        listas1[1].append(data[0].firstChild.data)
        listas1[2].append(nome[0].firstChild.data)
        listas1[3].append(cidade[1].firstChild.data)
        listas1[4].append(UF[1].firstChild.data)
        listas1[5].append(loja[0].firstChild.data)
        listas1[6].append(NF[0].firstChild.data)
        listas1[7].append(len(itens))
        listas1[8].append(prod[-1].firstChild.data)
        listas1[9].append(frete[-1].firstChild.data)
        listas1[10].append(desc[-1].firstChild.data)
        listas1[11].append(total[0].firstChild.data)

cod = []
for item in listas1[5]:                         
    sliced = item.split("Loja: ")[::-1]
    codigo = sliced[0]
    if codigo[0:4] == "Loja":
        cod.append("Americanas")
    elif codigo[0:4] == "Subm":
        cod.append("Americanas")
    elif codigo[0:4] == "Shop":
        cod.append("Americanas")
    elif codigo[0] == "2":
        cod.append("Shopee")
    elif codigo[0] == "9":
        cod.append("Netshoes")
    elif codigo[0:4] == "2000":
        cod.append("Integra")
    elif codigo[0] == "7":
        cod.append("Amazon")
    elif codigo[3:4] == "11":
        cod.append("Integra")
    else:
        cod.append("0")
        
listas1.pop(5)
listas1.insert(5,cod)
fill(listas1,cab_1)
fill(listas2,cab_2)

nfe_dict = {x[0]: x[1:] for x in listas1}
df = pd.DataFrame(nfe_dict)
df.columns = headers

df['Data'] = pd.to_datetime(df.Data)
df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')
for col in df.columns[8:]:
    df[col] = pd.to_numeric(df[col], errors='coerce')
df['Número da NF'] = df['Número da NF'].astype(str)
df['Itens comprados'] = df['Itens comprados'].astype(str)

df.to_excel()

C2 = t.time()
exec_time(C1,C2)
