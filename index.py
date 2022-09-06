from random import randint
from docxtpl import DocxTemplate,InlineImage

doc = DocxTemplate('Template/Document_template.docx')

sales_row = []

for iter in range(10):
        costpu = randint(1,15)
        nUnits = randint(100,500)
        sales_row.append({'SrNo':iter+1,'name':'item'+str(iter+1),
        'costpu':costpu,'nUnits':nUnits, 'revenue':costpu*nUnits})


context = {
"salesTblRows":sales_row,
"Chart":InlineImage(doc,'Images/line.png')
}

doc.render(context)

doc.save('Output/report.docx')
