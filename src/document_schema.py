"""
Document field mapping based on RESIDENTIAL_Checks.xlsx
Defines what fields to extract from each document type
"""

# Based on "Checks" tab - correspondence matrix
DOCUMENT_FIELDS = {
    'CONTRATO_CESION_AHORROS': {
        'sd_do_company_name': None,
        'sd_do_cif': None,
        'sd_do_address': None,
        'sd_do_representative_name': None,
        'sd_do_representative_dni': None,
        'sd_do_signature': None,
        'homeowner_name': None,
        'homeowner_dni': None,
        'homeowner_address': None,
        'homeowner_phone': None,
        'homeowner_email': None,
        'homeowner_notifications': None,
        'homeowner_signature': None,
    },
    
    'FICHA_RES020': {
        'act_code': None,  # 010/020
        'energy_savings_kwh': None,
        'start_date': None,
        'finish_date': None,
        'address': None,
        'catastral_ref': None,
        'utm_coordinates': None,
        'climatic_zone': None,
        'investment_eur': None,
        'remuneration_eur_kwh': None,
        'lifespan_years': None,
        'sell_price_eur_kwh': None,
        'surface_measurements': None,
        'photos': None,
    },
    
    'DECLARACION_RESPONSABLE': {
        # Fields from Checks tab
    },
    
    'FACTURA': {
        'invoice_number': None,
        'invoice_date': None,
        # More fields
    },
    
    'INFORME_FOTOGRAFICO': {
        'photos_before': None,
        'photos_after': None,
        'measurements': None,
        'house_photo': None,
        'catastro_image': None,
    },
    
    'CERTIFICADO_INSTALADOR': {
        'installer_name': None,
        'installer_address': None,
        'installer_cif': None,
        'installer_signature': None,
    },
    
    'CEE_FINAL': {
        'cee_date': None,
        'cee_signature': None,
    },
    
    'REGISTRO_CEE': {
        'registration_date': None,
        'registration_number': None,
    },
    
    'DNI': {
        'dni_number': None,
        'dni_name': None,
    },
}