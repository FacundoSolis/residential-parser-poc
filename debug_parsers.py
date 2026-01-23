"""
Debug script to test parsers individually
"""
import sys
from pathlib import Path

# Fix imports
sys.path.insert(0, str(Path(__file__).parent))

from src.matrix_generator import MatrixGenerator

# Ruta al ZIP/carpeta del cliente
folder = Path("data/ALEGRE SANTA CRUZ MARIA PILAR")

print("="*80)
print("🐛 DEBUG: Parsing individual documents")
print("="*80)

# Crear generator
mg = MatrixGenerator(str(folder))
mg.parse_all_documents()

print("\n" + "="*80)
print("📊 PARSED DATA SUMMARY")
print("="*80)

# --- CONTRATO ---
print("\n🔹 CONTRATO:")
if 'CONTRATO' in mg.parsed_data:
    contrato = mg.parsed_data['CONTRATO']
    print(f"  catastral_ref: '{contrato.get('catastral_ref', '')}'")
    print(f"  lifespan: '{contrato.get('lifespan', '')}'")
    print(f"  act_code: '{contrato.get('act_code', '')}'")
else:
    print("  ❌ NOT FOUND")

# --- CERTIFICADO ---
print("\n🔹 CERTIFICADO:")
if 'CERTIFICADO' in mg.parsed_data:
    cert = mg.parsed_data['CERTIFICADO']
    print(f"  start_date: '{cert.get('start_date', '')}'")
    print(f"  finish_date: '{cert.get('finish_date', '')}'")
    print(f"  lifespan: '{cert.get('lifespan', '')}'")
    print(f"  catastral_ref: '{cert.get('catastral_ref', '')}'")
    print(f"  b: '{cert.get('b', '')}'")
    print(f"  calculation_methodology: '{cert.get('calculation_methodology', '')}'")
else:
    print("  ❌ NOT FOUND")

# --- CEE ---
print("\n🔹 CEE:")
if 'CEE' in mg.parsed_data:
    cee = mg.parsed_data['CEE']
    print(f"  address: '{cee.get('address', '')}'")
    print(f"  catastral_ref: '{cee.get('catastral_ref', '')}'")
    print(f"  certification_date: '{cee.get('certification_date', '')}'")
else:
    print("  ❌ NOT FOUND")

# --- CALCULO ---
print("\n🔹 CALCULO:")
if 'CALCULO' in mg.parsed_data:
    calc = mg.parsed_data['CALCULO']
    print(f"  zone_climatique (G): '{calc.get('zone_climatique', '')}'")
    print(f"  calculation_methodology (0,83): '{calc.get('calculation_methodology', '')}'")
    print(f"  act_code: '{calc.get('act_code', '')}'")
else:
    print("  ❌ NOT FOUND")

# --- FICHA ---
print("\n🔹 FICHA:")
if 'FICHA' in mg.parsed_data:
    ficha = mg.parsed_data['FICHA']
    print(f"  act_code: '{ficha.get('act_code', '')}'")
else:
    print("  ❌ NOT FOUND")

print("\n" + "="*80)
print("🧪 TESTING _pick_best for Lifespan")
print("="*80)

cert_life = mg._get_value('CERTIFICADO', 'lifespan')
contr_life = mg._get_value('CONTRATO', 'lifespan')
life_result = mg._pick_best(cert_life, contr_life, min_len=1)

print(f"  CERTIFICADO.lifespan = '{cert_life}'")
print(f"  CONTRATO.lifespan = '{contr_life}'")
print(f"  _pick_best result = '{life_result}'")
print(f"  _is_valid_cell_value(cert_life) = {mg._is_valid_cell_value(cert_life, min_len=1)}")

print("\n" + "="*80)
print("🧪 TESTING _pick_best for Catastral ref")
print("="*80)

cert_cat = mg._get_value('CERTIFICADO', 'catastral_ref')
cee_cat = mg._get_value('CEE', 'catastral_ref')
reg_cat = mg._get_value('REGISTRO', 'catastral_ref')
contr_cat = mg._get_value('CONTRATO', 'catastral_ref')
decl_cat = mg._get_value('DECLARACION', 'catastral_ref')

cat_result = mg._pick_best(cert_cat, cee_cat, reg_cat, contr_cat, decl_cat, min_len=10)

print(f"  CERTIFICADO.catastral_ref = '{cert_cat}'")
print(f"  CEE.catastral_ref = '{cee_cat}'")
print(f"  REGISTRO.catastral_ref = '{reg_cat}'")
print(f"  CONTRATO.catastral_ref = '{contr_cat}'")
print(f"  DECLARACION.catastral_ref = '{decl_cat}'")
print(f"  _pick_best result = '{cat_result}'")

print("\n" + "="*80)
print("✅ DEBUG COMPLETE")
print("="*80)