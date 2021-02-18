from os import listdir
from tika import parser
from openpyxl import load_workbook
import time


start_time = time.time()
list_of_values=[[]]


def get_piece(text,start,end):
        if text == "" or start == "" or end == "":
            print("No pdf text or no start/end mapping for required field")
            return
        else:
            return text.split(start)[1].split(end)[0]
    

mypathpdf="C:\\python\\Projects\\EM2P\\PDF files\\"
all_pdf_files=listdir(mypathpdf)

workbook = load_workbook(filename="list of values.xlsx")
sheet = workbook.active

for each_pdf in all_pdf_files:
        
    file = mypathpdf+each_pdf
    file_data = parser.from_file(file)
    text = file_data['content']
   
    ####Extract Nº contrato
    ##in_between
    start = "Cliente: Nº Contrato: "
    end = "\n\nDADOS DO CONTRATO"
    contrato = get_piece(text,start,end)
    
    ####Extract periodo de faturação
    #to_the_lef_fixed
    match="Período de faturação\n\nTotal Faturado no mês"
    ind_per_fat=text.find(match)
    
    #len("dd/mm/yyyy") = 10
    #len("27/10/2020 a 26/11/2020") = 23 and assumed to be fixed
    #to_the_lef_fixed
    #if not we are in trouble...
    periodo_fat_init = text[ind_per_fat-23:ind_per_fat-(23-10)] 
    periodo_fat_fin = text[ind_per_fat-10:ind_per_fat-1] 
    
    #####Extract num factura
    #VNFACR/200202111774\n\n
    match = "DATA EMISSÃOVNFACR/"
    ind_start = text.find(match)
    ind_end = text.find("\n\n",ind_start)
    num_fatura = text[ind_start+12:ind_end] #len("DATA EMISSÃO") = 12
    
    #####Extract data emissão
    match = "DATA EMISSÃO\n\n"
    ind_start = text.find(match)
    data_emissao = text[ind_start-10:ind_start]
    
    #####Extract valor a pagar
    match = "€\n\nValor a Pagar\n\n"
    ind_end = text.find(match)
    match2 = "Data Limite Pagamento**\n\n" #len = 25
    ind_start = text.find(match2)+25
    valor_a_pagar = text[ind_start:ind_end]
    
    
    #####Extract total faturado no mês 
    #to_the_right_not_fixed
    match = "Total Faturado no mês\n\n" #len=23
    match2 = "\n\nData Limite Pagamento**\n\n"
    ind_start = text.find(match) + 23
    ind_end = text.find(match2)-1
    total_faturado_mes = text[ind_start:ind_end]
    
    print(f"Nº contrato - {contrato} Faturado de {periodo_fat_init} a"
          f" {periodo_fat_fin} \nNº fatura {num_fatura} Emitida a"
          f" {data_emissao} Valor a pagar - {valor_a_pagar} €"
          f" Total faturado do mês - {total_faturado_mes}")
    
    
    list_of_values=[contrato,periodo_fat_init,periodo_fat_fin,num_fatura,data_emissao,valor_a_pagar,total_faturado_mes]
    
    
    # text_file = open(each_pdf.replace("pdf","txt"), "w")
    # text_file.write(text)
    # text_file.close()
       
    sheet.append(list_of_values)
    
workbook.save(filename="list of values.xlsx")
    
print("--- %s seconds ---" % (time.time() - start_time))
