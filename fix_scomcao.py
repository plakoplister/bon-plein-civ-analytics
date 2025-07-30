import json

# Charger les données
with open('WEBAPP_PUBLICATION/dynamic_data_enriched.json', 'r') as f:
    data = json.load(f)

# Compter les occurrences SCOMCAO
scomcao_count = sum(1 for record in data['records'] if record.get('exportateur_simple') == 'SCOMCAO')
print(f'Transactions SCOMCAO trouvées: {scomcao_count}')

# Remplacer SCOMCAO par S3C
for record in data['records']:
    if record.get('exportateur_simple') == 'SCOMCAO':
        record['exportateur_simple'] = 'S3C'
    if record.get('exportateur') and 'SCOMCAO' in record['exportateur']:
        record['exportateur'] = record['exportateur'].replace('SCOMCAO', 'S3C')

# Mettre à jour les filtres
if 'SCOMCAO' in data['filters']['exportateurs']:
    new_exportateurs = []
    for exp in data['filters']['exportateurs']:
        if exp == 'SCOMCAO':
            new_exportateurs.append('S3C')
        else:
            new_exportateurs.append(exp)
    # Enlever les doublons et trier
    data['filters']['exportateurs'] = sorted(list(set(new_exportateurs)))

# Sauvegarder
with open('WEBAPP_PUBLICATION/dynamic_data_enriched.json', 'w') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

print(f'✅ {scomcao_count} transactions SCOMCAO converties en S3C')
print('✅ Filtres mis à jour')