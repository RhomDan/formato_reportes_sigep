import io
from flask import Flask, render_template, request, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)

def formato_consistencia_conta_teso_grupo(nombre_archivo):
        try :
                df = pd.read_excel(nombre_archivo, 
                                skiprows = 16, 
                                skipfooter = 1, 
                                header = None,
                                dtype = object,
                                engine = 'xlrd')
        except:
                nombre_archivo.seek(0)
                df = pd.read_excel(nombre_archivo, 
                                skiprows = 16, 
                                skipfooter = 1, 
                                header = None,
                                dtype = object,
                                engine = 'xlrd')
        cols = ['tipo', 'nro_asiento', 'entidad_conta', 'da_conta', 'comprobante', 'tipo', 'regularizacion','transferencia','monto_conta',
                'entidad_teso', 'da_teso', 'comprobante_teso','t_teso', 'tipo_teso', 'codigo_op', 'monto_teso', 'diferencia']
        df = df.dropna(axis = 1, how = 'all')
        df.columns = cols
        conta = df.iloc[:,0:9].copy()
        conta = conta.dropna(axis = 0)
        teso = df.iloc[:,9:-1].copy()
        conta['key'] = 'Entidad:'+ conta['entidad_conta'].astype(str) + 'DA:' +conta['da_conta'].astype(str) + 'Comprobante:' + conta['comprobante']
        teso['key'] = 'Entidad:'+ teso['entidad_teso'].astype(str) + 'DA:' +teso['da_teso'].astype(str) + 'Comprobante:' + teso['comprobante_teso']
        contraste = conta.merge(right = teso, how = 'left', on = 'key')
        output = io.BytesIO()
        return archivo_final(output, contraste)

def archivo_final(salida, df):
        with pd.ExcelWriter(salida, engine = 'openpyxl') as writer:
                df.to_excel(writer, index = False, sheet_name = 'Procesados')
                worksheet = writer.sheets['Procesados']
                header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
                header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') # Azul oscuro            
                center_alignment = Alignment(horizontal='center', vertical='center')
                for col_num in range(1, len(df.columns) + 1):
                        cell = worksheet.cell(row=1, column=col_num)
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = center_alignment
                for col in worksheet.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = get_column_letter(col[0].column)
                        worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
        salida.seek(0)
        return salida

tipos_reportes = {
        'analisis_consistencia':formato_consistencia_conta_teso_grupo
}

@app.route('/')
def index():
        return render_template('index.html')
@app.route('/transform', methods = ['POST'])
def transform():
        report_type = request.form.get('report_type')
        file = request.files.get('excel_file')
        if file.filename == '':
                return 'Nombre de archivo no valido', 400
        if report_type not in tipos_reportes:
                return "Tipo de reporte no válido.", 400
        funcion_procesadora = tipos_reportes[report_type]
        processed_file = funcion_procesadora(file)
        output_filename = f"procesado_{file.filename.rsplit('.', 1)[0]}.xlsx"
        return send_file(
                        processed_file,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        as_attachment=True,
                        download_name=output_filename
                )
if __name__ == '__main__':
        app.run(debug = True)