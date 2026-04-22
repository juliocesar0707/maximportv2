import re

with open('schema_rrv_estoque.txt', 'r', encoding='utf-8') as f:
    content = f.read()

target_tables = ['produto', 'produto_empresa', 'cliente', 'financeiro', 'proncm', 'produtoUn', 'prolote']

with open('target_tables.md', 'w', encoding='utf-8') as f:
    for t in target_tables:
        pattern = rf'## Tabela: {t}\n.*?(?=\n## Tabela: |\Z)'
        match = re.search(pattern, content, re.DOTALL)
        if match:
            f.write(match.group(0) + '\n\n')
        else:
            f.write(f"## Tabela: {t}\nNÃO ENCONTRADA\n\n")
