"""
Document field mapping based on RESIDENTIAL_Checks.xlsx
Defines what fields to extract from each document type
"""

# Based on "Checks" tab - correspondence matrix
DOCUMENT_FIELDS = {
    'CONTRATO_CESION_AHORROS': [
        'sd_do_company_name',
        'sd_do_cif',
        'sd_do_address',
        'sd_do_representative_name',
        'sd_do_representative_dni',
        'sd_do_signature',
        'homeowner_name',
        'homeowner_dni',
        'homeowner_address',
        'homeowner_phone',
        'homeowner_email',
        'homeowner_notifications',
        'homeowner_signatures',
    ],
    
    'FICHA_RES020': [
        'act_code',
        'energy_savings_kwh',
        'start_date',
        'finish_date',
        'address',
        'catastral_ref',
        'utm_coordinates',
        'climatic_zone',
        'investment_eur',
        'remuneration_eur_kwh',
        'lifespan_years',
        'sell_price_eur_kwh',
        'calculation_values',
        'surface',
        'photos',
    ],
    
    'DECLARACION_RESPONSABLE': [
        'representative_name',
        'representative_dni',
        'signature',
    ],
    
    'FACTURA': [
        'invoice_number',
        'invoice_date',
    ],
    
    'INFORME_FOTOGRAFICO': [
        'photos_before_after',
        'measurements',
        'house_photo',
        'catastro_image',
    ],
    
    'CERTIFICADO_INSTALADOR': [
        'installer_name',
        'installer_address',
        'installer_cif',
        'installer_signature',
    ],
    
    'CEE_FINAL': [
        'cee_signature',
        'cee_date',
    ],
    
    'REGISTRO_CEE': [
        'registration_date',
    ],
    
    'DNI': [
        'dni_number',
        'dni_name',
    ],
}

def get_document_type(filename: str) -> str:
    """Identify document type from filename"""
    filename_upper = filename.upper()
    
    if 'CONTRATO' in filename_upper:
        return 'CONTRATO_CESION_AHORROS'
    elif 'FICHA' in filename_upper:
        return 'FICHA_RES020'
    elif 'DECLARACION' in filename_upper:
        return 'DECLARACION_RESPONSABLE'
    elif 'FACTURA' in filename_upper:
        return 'FACTURA'
    elif 'FOTOGRAFICO' in filename_upper or 'FOTOGRÁFICO' in filename_upper:
        return 'INFORME_FOTOGRAFICO'
    elif 'CERTIFICADO' in filename_upper:
        return 'CERTIFICADO_INSTALADOR'
    elif 'CEE FINAL' in filename_upper:
        return 'CEE_FINAL'
    elif 'REGISTRO' in filename_upper:
        return 'REGISTRO_CEE'
    elif 'DNI' in filename_upper:
        return 'DNI'
    else:
        return 'UNKNOWN'
