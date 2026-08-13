import pandas as pd
from pathlib import Path
import subprocess
import os
import re

REPO = Path(__file__).parent
UPLOAD = REPO / "upload"

anos = [2022, 2023, 2024, 2025, 2026]

arquivos_gerados = 0


def anonimizar_cpf(valor):
    """
    Anonimiza CPF mantendo o formato XXX.XXX.XXX-XX, mascarando o
    primeiro e o último bloco: ***.XXX.XXX-**

    Regras:
    - CNPJ (14 dígitos) não é alterado, pois identifica pessoa jurídica.
    - CPF pode chegar sem os zeros à esquerda (ex: 5190403688, com 10
      dígitos, na verdade é 051.904.036-88). Nesse caso completa com
      zeros à esquerda até 11 dígitos antes de aplicar a máscara.
    - Valores vazios/nulos são retornados sem alteração.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return valor

    texto = str(valor).strip()
    if not texto:
        return valor

    digitos = re.sub(r"\D", "", texto)

    if not digitos:
        return valor

    # CNPJ (pessoa jurídica) -> não anonimiza
    if len(digitos) == 14:
        return valor

    # CPF -> completa com zeros à esquerda até 11 dígitos
    if len(digitos) < 11:
        digitos = digitos.zfill(11)
    elif len(digitos) > 11:
        # segurança: caso venha com dígitos a mais, mantém os 11 finais
        digitos = digitos[-11:]

    return f"***.{digitos[3:6]}.{digitos[6:9]}-**"


def anonimizar_coluna_cpf(df):
    """Localiza a coluna de CPF/CNPJ do credor (se existir) e anonimiza
    apenas os valores de CPF, mantendo CNPJ intacto."""
    for col in df.columns:
        if "cpf" in str(col).lower():
            df[col] = df[col].apply(anonimizar_cpf)
            print(f"Coluna anonimizada: {col}")
    return df



for ano in anos:

    arquivo = UPLOAD / f"DESPESAPOREMPENHO{ano}.xlsx"

    if not arquivo.exists():
        print(f"Arquivo não encontrado: {arquivo.name}")
        continue

    print(f"Processando {arquivo.name}...")

    abas = pd.read_excel(
        arquivo,
        sheet_name=None,
        header=1
    )

    for aba in ["orgão", "empenho"]:

        if aba not in abas:
            print(f"Aba '{aba}' não encontrada em {arquivo.name}")
            continue

        df = abas[aba]

        # remove primeira coluna
        df = df.iloc[:, 1:]

        # remove linhas vazias
        df = df.dropna(how="all")

        # remove colunas vazias
        df = df.dropna(axis=1, how="all")

        # anonimiza CPF (CNPJ permanece intacto)
        df = anonimizar_coluna_cpf(df)

        saida = UPLOAD / f"{aba}{ano}.xlsx"

        df.to_excel(
            saida,
            index=False
        )

        print(f"Gerado: {saida.name}")

        arquivos_gerados += 1

    os.remove(arquivo)
    print(f"Removido: {arquivo.name}")

# ==========================
# GIT
# ==========================

if arquivos_gerados > 0:

    resultado = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=REPO,
        capture_output=True,
        text=True
    )

    if resultado.stdout.strip():

        subprocess.run(
            ["git", "add", "upload"],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            [
                "git",
                "commit",
                "-m",
                "Atualização automática portal despesa"
            ],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            ["git", "push"],
            cwd=REPO,
            check=True
        )

        print("GitHub atualizado com sucesso.")

    else:
        print("Nenhuma alteração encontrada.")

else:
    print("Nenhum arquivo foi gerado.")