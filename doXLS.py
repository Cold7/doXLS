import xlsxwriter
from glob import glob
from math import fabs


def doXLS(deseq,edger, xlsName):
	#parsing gene names
	geneNames=open("c_elegans.PRJNA13758.WS256.geneIDs.txt","r")
	geneNameDict={}
	for line in geneNames:
		if "Live" in line:
			splitted= line[:-1].split(",")
			geneNameDict[splitted[1]]=[splitted[2],splitted[3]]

	#creating xls
	workbook = xlsxwriter.Workbook(xlsName+".xls")
	filesDeseq=glob(deseq+"/*")
	filesEdgeR=glob(edger+"/*")
	
	#looking for the corresponding file in both folders
	for fileD in filesDeseq:
		nameD=fileD.split(deseq+"/")[1][:-4]
		for fileE in filesEdgeR:
			nameE=fileE.split(edger+"/")[1][:-4]
			if nameD == nameE:
				listD=[[],[],[]] #[gene name] [log] [padj]
				listE= [[],[],[]] ##[gene name] [log] [fdr]
				#creating worksheet
				worksheet = workbook.add_worksheet(nameD)
				
				#######################################
				#parsing deseq2 and edger results
				#######################################
				
				#opening first file from deseq2
				DGED=open(fileD,"r")
				flag=0 #to skip the first line
				for line in DGED:
					if flag!=0:
						splitted = line.split(",")
						try:
							if fabs(float(splitted[3]))>=2 and float(splitted[7])<=0.05:
								listD[0].append(splitted[1][1:-1])
								listD[1].append(float(splitted[3]))
								listD[2].append(float(splitted[7]))
						except:
							pass
					flag=1
				
				DGED.close()

				#opening second file from edger
				DGEE=open(fileE,"r")
				flag=0 #to skip the first line
				for line in DGEE:
					if flag!=0:
						splitted = line.split("\t")
						try:
							if fabs(float(splitted[1]))>=2 and float(splitted[5])<=0.05:
								listE[0].append(splitted[0])
								listE[1].append(float(splitted[1]))
								listE[2].append(float(splitted[5]))
						except:
							pass
					flag=1
				
				DGEE.close()	
				
				#so, the last step is to put all information toghether
				#first we will put all genes that are DE according to EdgeR and DESeq2,
				#then only  in Deseq2 and at least only in edgeR

				#Adding names to cols
				worksheet.write(0,0, 'DE found using both software')
				worksheet.write(0,1, 'Name')
				worksheet.write(0,2, 'Name')
				worksheet.write(0,3, 'log2 DESeq2')
				worksheet.write(0,4, 'padj DESeq2')
				worksheet.write(0,5, 'log2 edgeR')
				worksheet.write(0,6, 'FDR edgeR')

				worksheet.write(0,8, 'DE found only through DESeq2')
				worksheet.write(0,9, 'Name')
				worksheet.write(0,10, 'Name')				
				worksheet.write(0,11, 'log2 DESeq2')
				worksheet.write(0,12, 'padj DESeq2')

				worksheet.write(0,14, 'DE found only through edgeR')
				worksheet.write(0,15, 'Name')
				worksheet.write(0,16, 'Name')				
				worksheet.write(0,17, 'log2 edgeR')
				worksheet.write(0,18, 'FDR edgeR')		
				
				#cont 1,2 and 3 are to know in wich position to write in the xls file
				cont0=1
				cont1=1
				cont2=1	
				for i in range(len(listD[0])):
					flag=0
					for j in range(len(listE[0])):
						if listD[0][i]==listE[0][j]:
							flag=1
							worksheet.write(cont0,0, listD[0][i])
							try:
								worksheet.write(cont0,1, geneNameDict[listD[0][i]][0])
								worksheet.write(cont0,2, geneNameDict[listD[0][i]][1])
							except:
								worksheet.write(cont0,1, "")
								worksheet.write(cont0,2, "")
							worksheet.write(cont0,3, float(listD[1][i]))
							worksheet.write(cont0,4, float(listD[2][i]))
							worksheet.write(cont0,5, float(listE[1][j]))
							worksheet.write(cont0,6, float(listE[2][j]))
							cont0+=1
					if flag==0:
							worksheet.write(cont1,8, listD[0][i])
							try:
								worksheet.write(cont1,9, geneNameDict[listD[0][i]][0])
								worksheet.write(cont1,10, geneNameDict[listD[0][i]][1])
							except:
								worksheet.write(cont1,9, "")
								worksheet.write(cont1,10, "")							
							worksheet.write(cont1,11, float(listD[1][i]))
							worksheet.write(cont1,12, float(listD[2][i]))
							cont1+=1
				
				for i in range(len(listE[0])):
					flag=0
					for j in range(len(listD[0])):
						if listD[0][j]==listE[0][i]:
							flag=1
					if flag==0:
							worksheet.write(cont2,14, listE[0][i])
							try:
								worksheet.write(cont2,15, geneNameDict[listE[0][i]][0])
								worksheet.write(cont2,16, geneNameDict[listE[0][i]][1])
							except:
								worksheet.write(cont2,15, "")
								worksheet.write(cont2,16, "")							
							worksheet.write(cont2,17, float(listE[1][i]))
							worksheet.write(cont2,18, float(listE[2][i]))						
							cont2+=1

	workbook.close()


if __name__=="__main__":
	doXLS("./DESeq2","./edgeR","output")
