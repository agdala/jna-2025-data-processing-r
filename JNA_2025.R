library(tidyverse)
library(readxl)
library(writexl)
library(utf8)
library(lookup)

#load dataset
#banco_dados_jna_2025 <- read_excel ("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_2025_raw_data_codebook.xlsx")
#this years JNA 2025 data analysis was conducted in the Power Bi Dashboard (available in Brazil - Information Management)

banco_dados_jna_2025<-banco_dados_jna_2025 %>%
  rename( data_inicio=data_inicio,
         data_fim=data_fim,
         duracao=duracao,
          base_contatos=v1,
         ja_participou_pesquisa=v2,
         entrevistado_venezuelano_a=v3,
         maior_idade=v4,
         chefe_hogar=v5,
         saber_responder=v6,
         entrevistado_genero=v7,
         municipio=v8,
         estado=v9,
         ano_chegada_brasil=v10,
         interiorizado_a=v11,
         intencao_movimento=v12,
         tamanho_hogar=v13,
         entrevistado_parentesco_chefe_hogar=v14,
         entrevistado_idade=v15,
         entrevistado_gravida_lactante=v16,
         entrevistado_gravida_servico_nutricao=v17,
         entrevistado_etnia=v18,
         entrevistado_deficiencia=v19,
         entrevistado_deficiencia_fisica=v20,
         entrevistado_deficiencia_mental=v21,
         entrevistado_deficiencia_auditiva=v22,
         entrevistado_deficiencia_visual=v23,
         entrevistado_deficiencia_autismo=v24,
         entrevistado_deficiencia_outro=v25,
         entrevistado_documento_refugio_residencia=v26,
         entrevistado_cpf=v27,
         entrevistado_cnh=v28,
         entrevistado_sus=v29,
         entrevistado_escola=v30,
         entrevistado_barreira_escola=v31,
         entrevistado_grau_escolar=v32,
         entrevistado_grau_escolar_completo=v33,
         entrevistado_trabalhando=v34,
         entrevistado_contrato_trabalho=v35,
         entrevistado_microempresa=v36,
         pessoa1_genero=v37_1,
         pessoa1_parentesco=v39_1,
         pessoa1_idade_anos=v40_1_o1,
         pessoa1_idade_meses=v40_1_o2,
         pessoa1_gravida_lactante=v41_1,
         pessoa1_gravida_servico_nutricao=v42_1,
         pessoa1_servico_nutricao_criancas=v43_1,
         pessoa1_etnia=v44_1,
         pessoa1_deficiencia=v45_1,
         pessoa1_deficiencia_fisica=v46_1,
         pessoa1_deficiencia_mental=v47_1,
         pessoa1_deficiencia_auditiva=v48_1,
         pessoa1_deficiencia_visual=v49_1,
         pessoa1_deficiencia_autismo=v50_1,
         pessoa1_deficiencia_outro=v51_1,
         pessoa1_nacionalidade=v52_1,
         pessoa1_documento_refugio_residencia=v53_1,
         pessoa1_cpf=v54_1,
         pessoa1_cnh=v55_1,
         pessoa1_certidao_nascimento=v56_1,
         pessoa1_sus=v57_1,
         pessoa1_escola=v58_1,
         pessoa1_barreira_escola=v59_1,
         pessoa1_grau_escolar=v60_1,
         pessoa1_grau_escolar_completo=v61_1,
         pessoa1_trabalhando=v62_1,
         pessoa1_contrato_trabalho=v63_1,
         pessoa1_microempresa=v64_1,
         pessoa2_genero=v37_2,
         pessoa2_parentesco=v39_2,
         pessoa2_idade_anos=v40_2_o1,
         pessoa2_idade_meses=v40_2_o2,
         pessoa2_gravida_lactante=v41_2,
         pessoa2_gravida_servico_nutricao=v42_2,
         pessoa2_servico_nutricao_criancas=v43_2,
         pessoa2_etnia=v44_2,
         pessoa2_deficiencia=v45_2,
         pessoa2_deficiencia_fisica=v46_2,
         pessoa2_deficiencia_mental=v47_2,
         pessoa2_deficiencia_auditiva=v48_2,
         pessoa2_deficiencia_visual=v49_2,
         pessoa2_deficiencia_autismo=v50_2,
         pessoa2_deficiencia_outro=v51_2,
         pessoa2_nacionalidade=v52_2,
         pessoa2_documento_refugio_residencia=v53_2,
         pessoa2_cpf=v54_2,
         pessoa2_cnh=v55_2,
         pessoa2_certidao_nascimento=v56_2,
         pessoa2_sus=v57_2,
         pessoa2_escola=v58_2,
         pessoa2_barreira_escola=v59_2,
         pessoa2_grau_escolar=v60_2,
         pessoa2_grau_escolar_completo=v61_2,
         pessoa2_trabalhando=v62_2,
         pessoa2_contrato_trabalho=v63_2,
         pessoa2_microempresa=v64_2,
         pessoa3_genero=v37_3,
         pessoa3_parentesco=v39_3,
         pessoa3_idade_anos=v40_3_o1,
         pessoa3_idade_meses=v40_3_o2,
         pessoa3_gravida_lactante=v41_3,
         pessoa3_gravida_servico_nutricao=v42_3,
         pessoa3_servico_nutricao_criancas=v43_3,
         pessoa3_etnia=v44_3,
         pessoa3_deficiencia=v45_3,
         pessoa3_deficiencia_fisica=v46_3,
         pessoa3_deficiencia_mental=v47_3,
         pessoa3_deficiencia_auditiva=v48_3,
         pessoa3_deficiencia_visual=v49_3,
         pessoa3_deficiencia_autismo=v50_3,
         pessoa3_deficiencia_outro=v51_3,
         pessoa3_nacionalidade=v52_3,
         pessoa3_documento_refugio_residencia=v53_3,
         pessoa3_cpf=v54_3,
         pessoa3_cnh=v55_3,
         pessoa3_certidao_nascimento=v56_3,
         pessoa3_sus=v57_3,
         pessoa3_escola=v58_3,
         pessoa3_barreira_escola=v59_3,
         pessoa3_grau_escolar=v60_3,
         pessoa3_grau_escolar_completo=v61_3,
         pessoa3_trabalhando=v62_3,
         pessoa3_contrato_trabalho=v63_3,
         pessoa3_microempresa=v64_3,
         pessoa4_genero=v37_4,
         pessoa4_parentesco=v39_4,
         pessoa4_idade_anos=v40_4_o1,
         pessoa4_idade_meses=v40_4_o2,
         pessoa4_gravida_lactante=v41_4,
         pessoa4_gravida_servico_nutricao=v42_4,
         pessoa4_servico_nutricao_criancas=v43_4,
         pessoa4_etnia=v44_4,
         pessoa4_deficiencia=v45_4,
         pessoa4_deficiencia_fisica=v46_4,
         pessoa4_deficiencia_mental=v47_4,
         pessoa4_deficiencia_auditiva=v48_4,
         pessoa4_deficiencia_visual=v49_4,
         pessoa4_deficiencia_autismo=v50_4,
         pessoa4_deficiencia_outro=v51_4,
         pessoa4_nacionalidade=v52_4,
         pessoa4_documento_refugio_residencia=v53_4,
         pessoa4_cpf=v54_4,
         pessoa4_cnh=v55_4,
         pessoa4_certidao_nascimento=v56_4,
         pessoa4_sus=v57_4,
         pessoa4_escola=v58_4,
         pessoa4_barreira_escola=v59_4,
         pessoa4_grau_escolar=v60_4,
         pessoa4_grau_escolar_completo=v61_4,
         pessoa4_trabalhando=v62_4,
         pessoa4_contrato_trabalho=v63_4,
         pessoa4_microempresa=v64_4,
         pessoa5_genero=v37_5,
         pessoa5_parentesco=v39_5,
         pessoa5_idade_anos=v40_5_o1,
         pessoa5_idade_meses=v40_5_o2,
         pessoa5_gravida_lactante=v41_5,
         pessoa5_gravida_servico_nutricao=v42_5,
         pessoa5_servico_nutricao_criancas=v43_5,
         pessoa5_etnia=v44_5,
         pessoa5_deficiencia=v45_5,
         pessoa5_deficiencia_fisica=v46_5,
         pessoa5_deficiencia_mental=v47_5,
         pessoa5_deficiencia_auditiva=v48_5,
         pessoa5_deficiencia_visual=v49_5,
         pessoa5_deficiencia_autismo=v50_5,
         pessoa5_deficiencia_outro=v51_5,
         pessoa5_nacionalidade=v52_5,
         pessoa5_documento_refugio_residencia=v53_5,
         pessoa5_cpf=v54_5,
         pessoa5_cnh=v55_5,
         pessoa5_certidao_nascimento=v56_5,
         pessoa5_sus=v57_5,
         pessoa5_escola=v58_5,
         pessoa5_barreira_escola=v59_5,
         pessoa5_grau_escolar=v60_5,
         pessoa5_grau_escolar_completo=v61_5,
         pessoa5_trabalhando=v62_5,
         pessoa5_contrato_trabalho=v63_5,
         pessoa5_microempresa=v64_5,
         pessoa6_genero=v37_6,
         pessoa6_parentesco=v39_6,
         pessoa6_idade_anos=v40_6_o1,
         pessoa6_idade_meses=v40_6_o2,
         pessoa6_gravida_lactante=v41_6,
         pessoa6_gravida_servico_nutricao=v42_6,
         pessoa6_servico_nutricao_criancas=v43_6,
         pessoa6_etnia=v44_6,
         pessoa6_deficiencia=v45_6,
         pessoa6_deficiencia_fisica=v46_6,
         pessoa6_deficiencia_mental=v47_6,
         pessoa6_deficiencia_auditiva=v48_6,
         pessoa6_deficiencia_visual=v49_6,
         pessoa6_deficiencia_autismo=v50_6,
         pessoa6_deficiencia_outro=v51_6,
         pessoa6_nacionalidade=v52_6,
         pessoa6_documento_refugio_residencia=v53_6,
         pessoa6_cpf=v54_6,
         pessoa6_cnh=v55_6,
         pessoa6_certidao_nascimento=v56_6,
         pessoa6_sus=v57_6,
         pessoa6_escola=v58_6,
         pessoa6_barreira_escola=v59_6,
         pessoa6_grau_escolar=v60_6,
         pessoa6_grau_escolar_completo=v61_6,
         pessoa6_trabalhando=v62_6,
         pessoa6_contrato_trabalho=v63_6,
         pessoa6_microempresa=v64_6,
         pessoa7_genero=v37_7,
         pessoa7_parentesco=v39_7,
         pessoa7_idade_anos=v40_7_o1,
         pessoa7_idade_meses=v40_7_o2,
         pessoa7_gravida_lactante=v41_7,
         pessoa7_gravida_servico_nutricao=v42_7,
         pessoa7_servico_nutricao_criancas=v43_7,
         pessoa7_etnia=v44_7,
         pessoa7_deficiencia=v45_7,
         pessoa7_deficiencia_fisica=v46_7,
         pessoa7_deficiencia_mental=v47_7,
         pessoa7_deficiencia_auditiva=v48_7,
         pessoa7_deficiencia_visual=v49_7,
         pessoa7_deficiencia_autismo=v50_7,
         pessoa7_deficiencia_outro=v51_7,
         pessoa7_nacionalidade=v52_7,
         pessoa7_documento_refugio_residencia=v53_7,
         pessoa7_cpf=v54_7,
         pessoa7_cnh=v55_7,
         pessoa7_certidao_nascimento=v56_7,
         pessoa7_sus=v57_7,
         pessoa7_escola=v58_7,
         pessoa7_barreira_escola=v59_7,
         pessoa7_grau_escolar=v60_7,
         pessoa7_grau_escolar_completo=v61_7,
         pessoa7_trabalhando=v62_7,
         pessoa7_contrato_trabalho=v63_7,
         pessoa7_microempresa=v64_7,
         pessoa8_genero=v37_8,
         pessoa8_parentesco=v39_8,
         pessoa8_idade_anos=v40_8_o1,
         pessoa8_idade_meses=v40_8_o2,
         pessoa8_gravida_lactante=v41_8,
         pessoa8_gravida_servico_nutricao=v42_8,
         pessoa8_servico_nutricao_criancas=v43_8,
         pessoa8_etnia=v44_8,
         pessoa8_deficiencia=v45_8,
         pessoa8_deficiencia_fisica=v46_8,
         pessoa8_deficiencia_mental=v47_8,
         pessoa8_deficiencia_auditiva=v48_8,
         pessoa8_deficiencia_visual=v49_8,
         pessoa8_deficiencia_autismo=v50_8,
         pessoa8_deficiencia_outro=v51_8,
         pessoa8_nacionalidade=v52_8,
         pessoa8_documento_refugio_residencia=v53_8,
         pessoa8_cpf=v54_8,
         pessoa8_cnh=v55_8,
         pessoa8_certidao_nascimento=v56_8,
         pessoa8_sus=v57_8,
         pessoa8_escola=v58_8,
         pessoa8_barreira_escola=v59_8,
         pessoa8_grau_escolar=v60_8,
         pessoa8_grau_escolar_completo=v61_8,
         pessoa8_trabalhando=v62_8,
         pessoa8_contrato_trabalho=v63_8,
         pessoa8_microempresa=v64_8,
         pessoa9_genero=v37_9,
         pessoa9_parentesco=v39_9,
         pessoa9_idade_anos=v40_9_o1,
         pessoa9_idade_meses=v40_9_o2,
         pessoa9_gravida_lactante=v41_9,
         pessoa9_gravida_servico_nutricao=v42_9,
         pessoa9_servico_nutricao_criancas=v43_9,
         pessoa9_etnia=v44_9,
         pessoa9_deficiencia=v45_9,
         pessoa9_deficiencia_fisica=v46_9,
         pessoa9_deficiencia_mental=v47_9,
         pessoa9_deficiencia_auditiva=v48_9,
         pessoa9_deficiencia_visual=v49_9,
         pessoa9_deficiencia_autismo=v50_9,
         pessoa9_deficiencia_outro=v51_9,
         pessoa9_nacionalidade=v52_9,
         pessoa9_documento_refugio_residencia=v53_9,
         pessoa9_cpf=v54_9,
         pessoa9_cnh=v55_9,
         pessoa9_certidao_nascimento=v56_9,
         pessoa9_sus=v57_9,
         pessoa9_escola=v58_9,
         pessoa9_barreira_escola=v59_9,
         pessoa9_grau_escolar=v60_9,
         pessoa9_grau_escolar_completo=v61_9,
         pessoa9_trabalhando=v62_9,
         pessoa9_contrato_trabalho=v63_9,
         pessoa9_microempresa=v64_9,
         pessoa10_genero=v37_10,
         pessoa10_parentesco=v39_10,
         pessoa10_idade_anos=v40_10_o1,
         pessoa10_idade_meses=v40_10_o2,
         pessoa10_gravida_lactante=v41_10,
         pessoa10_gravida_servico_nutricao=v42_10,
         pessoa10_servico_nutricao_criancas=v43_10,
         pessoa10_etnia=v44_10,
         pessoa10_deficiencia=v45_10,
         pessoa10_deficiencia_fisica=v46_10,
         pessoa10_deficiencia_mental=v47_10,
         pessoa10_deficiencia_auditiva=v48_10,
         pessoa10_deficiencia_visual=v49_10,
         pessoa10_deficiencia_autismo=v50_10,
         pessoa10_deficiencia_outro=v51_10,
         pessoa10_nacionalidade=v52_10,
         pessoa10_documento_refugio_residencia=v53_10,
         pessoa10_cpf=v54_10,
         pessoa10_cnh=v55_10,
         pessoa10_certidao_nascimento=v56_10,
         pessoa10_sus=v57_10,
         pessoa10_escola=v58_10,
         pessoa10_barreira_escola=v59_10,
         pessoa10_grau_escolar=v60_10,
         pessoa10_grau_escolar_completo=v61_10,
         pessoa10_trabalhando=v62_10,
         pessoa10_contrato_trabalho=v63_10,
         pessoa10_microempresa=v64_10,
         pessoa11_genero=v37_11,
         pessoa11_parentesco=v39_11,
         pessoa11_idade_anos=v40_11_o1,
         pessoa11_idade_meses=v40_11_o2,
         pessoa11_gravida_lactante=v41_11,
         pessoa11_gravida_servico_nutricao=v42_11,
         pessoa11_servico_nutricao_criancas=v43_11,
         pessoa11_etnia=v44_11,
         pessoa11_deficiencia=v45_11,
         pessoa11_deficiencia_fisica=v46_11,
         pessoa11_deficiencia_mental=v47_11,
         pessoa11_deficiencia_auditiva=v48_11,
         pessoa11_deficiencia_visual=v49_11,
         pessoa11_deficiencia_autismo=v50_11,
         pessoa11_deficiencia_outro=v51_11,
         pessoa11_nacionalidade=v52_11,
         pessoa11_documento_refugio_residencia=v53_11,
         pessoa11_cpf=v54_11,
         pessoa11_cnh=v55_11,
         pessoa11_certidao_nascimento=v56_11,
         pessoa11_sus=v57_11,
         pessoa11_escola=v58_11,
         pessoa11_barreira_escola=v59_11,
         pessoa11_grau_escolar=v60_11,
         pessoa11_grau_escolar_completo=v61_11,
         pessoa11_trabalhando=v62_11,
         pessoa11_contrato_trabalho=v63_11,
         pessoa11_microempresa=v64_11,
         pessoa12_genero=v37_12,
         pessoa12_parentesco=v39_12,
         pessoa12_idade_anos=v40_12_o1,
         pessoa12_idade_meses=v40_12_o2,
         pessoa12_gravida_lactante=v41_12,
         pessoa12_gravida_servico_nutricao=v42_12,
         pessoa12_servico_nutricao_criancas=v43_12,
         pessoa12_etnia=v44_12,
         pessoa12_deficiencia=v45_12,
         pessoa12_deficiencia_fisica=v46_12,
         pessoa12_deficiencia_mental=v47_12,
         pessoa12_deficiencia_auditiva=v48_12,
         pessoa12_deficiencia_visual=v49_12,
         pessoa12_deficiencia_autismo=v50_12,
         pessoa12_deficiencia_outro=v51_12,
         pessoa12_nacionalidade=v52_12,
         pessoa12_documento_refugio_residencia=v53_12,
         pessoa12_cpf=v54_12,
         pessoa12_cnh=v55_12,
         pessoa12_certidao_nascimento=v56_12,
         pessoa12_sus=v57_12,
         pessoa12_escola=v58_12,
         pessoa12_barreira_escola=v59_12,
         pessoa12_grau_escolar=v60_12,
         pessoa12_grau_escolar_completo=v61_12,
         pessoa12_trabalhando=v62_12,
         pessoa12_contrato_trabalho=v63_12,
         pessoa12_microempresa=v64_12,
         menores_18_hogar=v65,
         hogar_gasto_principal=v66,
         hogar_renda_mensal=v67,
         hogar_dificuldade_alimentacao=v68,
         hogar_razao_dificuldade_alimentacao=v69,
         hogar_necessidade_medica=v70,
         hogar_necessidade_medica_ginecologica=v71,
         hogar_necessidade_medica_pediatrica=v72,
         hogar_necessidade_medica_neurologica=v73,
         hogar_necessidade_medica_cardiologica=v74,
         hogar_necessidade_medica_saude_mental=v75,
         hogar_necessidade_medica_nutricional=v76,
         hogar_necessidade_medica_clinico_geral=v77,
         hogar_necessidade_medica_outro=v78,
         hogar_necessidade_medica_ns=v79,
         hogar_necessidade_medica_nr=v80,
         hogar_dificuldade_acesso_saude=v81,
         hogar_dificuldade_acesso_saude_lento=v82,
         hogar_dificuldade_acesso_saude_custo=v83,
         hogar_dificuldade_acesso_saude_documentacao=v84,
         hogar_dificuldade_acesso_saude_cheio=v85,
         hogar_dificuldade_acesso_saude_especialista=v86,
         hogar_dificuldade_acesso_saude_informacao=v87,
         hogar_dificuldade_acesso_saude_idioma=v88,
         hogar_dificuldade_acesso_saude_discriminacao=v89,
         hogar_dificuldade_acesso_saude_outro=v90,
         hogar_dificuldade_acesso_saude_ns=v91,
         hogar_dificuldade_acesso_saude_nr=v92,
         hogar_diarreia_disenteria=v93,
         hogar_acesso_agua=v94,
         hogar_saneamento=v95,
         hogar_coleta_lixo_semanal=v96,
         hogar_abrigo_utensilios_cozinha=v97,
         hogar_abrigo_cama=v98,
         hogar_abrigo_colchao=v99,
         hogar_abrigo_moveis=v100,
         hogar_abrigo_electrodomesticos=v101,
         hogar_abrigo_utensilios_nr=v102,
         hogar_moradia_proximo_mes=v103,
         hogar_ja_abrigo_antes=v104,
         hogar_discriminacao=v105,
         hogar_uasc=v106,
         hogar_trabalho_enganacao=v107,
         hogar_trabalho_sem_pagamento=v108,
         seguranca_vbg=v109)


#load reference tables for columns

v1 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v1")
v2 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v2")
v3 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v3")
v4 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v4")
v5 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v5")
v6 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v6")
v7_37 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v7_37")
v8 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v8")
v9 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v9")
v10 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v10")
v11 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v11")
v12 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v12")
v13 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v13")
v14_39 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v14_39")
v16_41 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v16_41")
v17_42_43 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v17_42_43")
v18 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v18")
v19_45 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v19_45")
v20_46 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v20_46")
v21_47 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v21_47")
v22_48 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v22_48")
v23_49 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v23_49")
v24_50 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v24_50")
v25 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v25")
v26_53 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v26_53")
v27_54 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v27_54")
v28_55 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v28_55")
v29_57 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v29_57")
v30_58 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v30_58")
v31 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v31")
v32_60 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v32_60")
v33_61 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v33_61")
v34 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v34")
v35_63 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v35_63")
v36_64 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v36_64")
v40 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v40")

v44_1 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_1")
v44_2 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_2")
v44_3 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_3")
v44_4 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_4")
v44_5 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_5")
v44_6 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_6")
v44_7 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_7")
v44_8 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_8")
v44_9 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_9")
v44_10 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_10")
v44_11 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_11")
v44_12 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v44_12")

v51_1 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v51_1")
v51_2 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v51_2")
v51_3 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v51_3")
v51_4 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v51_4")

v52_1 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_1")
v52_2 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_2")
v52_3 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_3")
v52_4 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_4")
v52_5 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_5")
v52_6 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_6")
v52_7 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_7")
v52_8 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_8")
v52_9 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_9")
v52_10 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_10")
v52_11 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_11")
v52_12 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v52_12")

v56 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v56")

v59_1 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_1")
v59_2 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_2")
v59_3 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_3")
v59_4 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_4")
v59_5 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_5")
v59_6 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_6")
v59_7 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_7")
v59_8 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_8")
v59_9 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_9")
v59_10 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_10")
v59_11 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_11")
v59_12 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v59_12")

v62_1 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_1")
v62_2 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_2")
v62_3 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_3")
v62_4 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_4")
v62_5 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_5")
v62_6 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_6")
v62_7 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_7")
v62_8 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_8")
v62_9 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_9")
v62_10 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_10")
v62_11 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_11")
v62_12 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v62_12")

v65 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v65")
v66 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v66")
v67 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v67")
v68 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v68")
v69 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v69")
v70 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v70")
v71 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v71")
v72 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v72")
v73 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v73")
v74 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v74")
v75 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v75")
v76 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v76")
v77 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v77")
v78 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v78")
v79 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v79")
v80 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v80")
v81 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v81")
v82 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v82")
v83 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v83")
v84 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v84")
v85 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v85")
v86 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v86")
v87 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v87")
v88 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v88")
v89 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v89")
v90 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v90")
v91 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v91")
v92 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v92")
v93 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v93")
v94 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v94")
v95 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v95")
v96 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v96")
v97 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v97")
v98 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v98")
v99 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v99")
v100 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v100")
v101 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v101")
v102 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v102")
v103 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v103")
v104 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v104")
v105 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v105")
v106 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v106")
v107 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v107")
v108 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v108")
v109 <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "v109")
idade_meses_ref <- read_excel("C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/jna_base_template.xlsx", sheet = "idade_meses_ref")


#relable responses, based on reference table

banco_dados_jna_2025$base_contatos = factor(banco_dados_jna_2025$base_contatos,
                                            levels = v1$levels,
                                            labels = v1$labels)

banco_dados_jna_2025$ja_participou_pesquisa = factor(banco_dados_jna_2025$ja_participou_pesquisa,
                                            levels = v2$levels,
                                            labels = v2$labels)

banco_dados_jna_2025$entrevistado_venezuelano_a = factor(banco_dados_jna_2025$entrevistado_venezuelano_a,
                                            levels = v3$levels,
                                            labels = v3$labels)

banco_dados_jna_2025$maior_idade = factor(banco_dados_jna_2025$maior_idade,
                                            levels = v4$levels,
                                            labels = v4$labels)

banco_dados_jna_2025$chefe_hogar = factor(banco_dados_jna_2025$chefe_hogar,
                                            levels = v5$levels,
                                            labels = v5$labels)

banco_dados_jna_2025$saber_responder = factor(banco_dados_jna_2025$saber_responder,
                                            levels = v6$levels,
                                            labels = v6$labels)

banco_dados_jna_2025$entrevistado_genero = factor(banco_dados_jna_2025$entrevistado_genero,
                                            levels = v7_37$levels,
                                            labels = v7_37$labels)

banco_dados_jna_2025$municipio = factor(banco_dados_jna_2025$municipio,
                                            levels = v8$levels,
                                            labels = v8$labels)

banco_dados_jna_2025$estado = factor(banco_dados_jna_2025$estado,
                                            levels = v9$levels,
                                            labels = v9$labels)

banco_dados_jna_2025$interiorizado_a = factor(banco_dados_jna_2025$interiorizado_a,
                                                 levels = v11$levels,
                                                 labels = v11$labels)

banco_dados_jna_2025$intencao_movimento = factor(banco_dados_jna_2025$intencao_movimento,
                                              levels = v12$levels,
                                              labels = v12$labels)


banco_dados_jna_2025$entrevistado_parentesco_chefe_hogar = factor(banco_dados_jna_2025$entrevistado_parentesco_chefe_hogar,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)

banco_dados_jna_2025$entrevistado_gravida_lactante = factor(banco_dados_jna_2025$entrevistado_gravida_lactante,
                                                     levels = v16_41$levels,
                                                     labels = v16_41$labels)

banco_dados_jna_2025$entrevistado_gravida_servico_nutricao = factor(banco_dados_jna_2025$entrevistado_gravida_servico_nutricao,
                                                            levels = v17_42_43$levels,
                                                            labels = v17_42_43$labels)

banco_dados_jna_2025$entrevistado_etnia= factor(banco_dados_jna_2025$entrevistado_etnia,
                                                                    levels = v18$levels,
                                                                    labels = v18$labels)


banco_dados_jna_2025$entrevistado_deficiencia= factor(banco_dados_jna_2025$entrevistado_deficiencia,
                                                levels = v19_45$levels,
                                                labels = v19_45$labels) 

banco_dados_jna_2025$entrevistado_deficiencia_fisica= factor(banco_dados_jna_2025$entrevistado_deficiencia_fisica,
                                                      levels = v20_46$levels,
                                                      labels = v20_46$labels) 

banco_dados_jna_2025$entrevistado_deficiencia_mental= factor(banco_dados_jna_2025$entrevistado_deficiencia_mental,
                                                             levels = v21_47$levels,
                                                             labels = v21_47$labels)

banco_dados_jna_2025$entrevistado_deficiencia_auditiva= factor(banco_dados_jna_2025$entrevistado_deficiencia_auditiva,
                                                             levels = v22_48$levels,
                                                             labels = v22_48$labels)

banco_dados_jna_2025$entrevistado_deficiencia_visual= factor(banco_dados_jna_2025$entrevistado_deficiencia_visual,
                                                               levels = v23_49$levels,
                                                               labels = v23_49$labels)

banco_dados_jna_2025$entrevistado_deficiencia_autismo= factor(banco_dados_jna_2025$entrevistado_deficiencia_autismo,
                                                             levels = v24_50$levels,
                                                             labels = v24_50$labels)

banco_dados_jna_2025$entrevistado_deficiencia_outro= factor(banco_dados_jna_2025$entrevistado_deficiencia_outro,
                                                              levels = v25$levels,
                                                              labels = v25$labels)

banco_dados_jna_2025$entrevistado_documento_refugio_residencia= factor(banco_dados_jna_2025$entrevistado_documento_refugio_residencia,
                                                            levels = v26_53$levels,
                                                            labels = v26_53$labels)

banco_dados_jna_2025$entrevistado_cpf= factor(banco_dados_jna_2025$entrevistado_cpf,
                                                            levels = v27_54$levels,
                                                            labels = v27_54$labels)

banco_dados_jna_2025$entrevistado_cnh= factor(banco_dados_jna_2025$entrevistado_cnh,
                                              levels = v28_55$levels,
                                              labels = v28_55$labels)

banco_dados_jna_2025$entrevistado_sus= factor(banco_dados_jna_2025$entrevistado_sus,
                                              levels = v29_57$levels,
                                              labels = v29_57$labels)

banco_dados_jna_2025$entrevistado_escola= factor(banco_dados_jna_2025$entrevistado_escola,
                                              levels = v30_58$levels,
                                              labels = v30_58$labels)

banco_dados_jna_2025$entrevistado_barreira_escola= factor(banco_dados_jna_2025$entrevistado_barreira_escola,
                                                 levels = v31$levels,
                                                 labels = v31$labels)

banco_dados_jna_2025$entrevistado_grau_escolar= factor(banco_dados_jna_2025$entrevistado_grau_escolar,
                                                          levels = v32_60$levels,
                                                          labels = v32_60$labels)

banco_dados_jna_2025$entrevistado_grau_escolar_completo= factor(banco_dados_jna_2025$entrevistado_grau_escolar_completo,
                                                       levels = v33_61$levels,
                                                       labels = v33_61$labels)

banco_dados_jna_2025$entrevistado_trabalhando= factor(banco_dados_jna_2025$entrevistado_trabalhando,
                                                                levels = v34$levels,
                                                                labels = v34$labels)

banco_dados_jna_2025$entrevistado_contrato_trabalho= factor(banco_dados_jna_2025$entrevistado_contrato_trabalho,
                                                      levels = v35_63$levels,
                                                      labels = v35_63$labels)

banco_dados_jna_2025$entrevistado_microempresa= factor(banco_dados_jna_2025$entrevistado_microempresa,
                                                            levels = v36_64$levels,
                                                            labels = v36_64$labels)

#pessoas idade
# 
# banco_dados_jna_2025$pessoa1_idade= factor(banco_dados_jna_2025$pessoa1_idade,
#                                                        levels = v40$levels,
#                                                        labels = v40$labels)
# banco_dados_jna_2025$pessoa2_idade= factor(banco_dados_jna_2025$pessoa2_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa3_idade= factor(banco_dados_jna_2025$pessoa3_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa4_idade= factor(banco_dados_jna_2025$pessoa4_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa5_idade= factor(banco_dados_jna_2025$pessoa5_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa6_idade= factor(banco_dados_jna_2025$pessoa6_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa7_idade= factor(banco_dados_jna_2025$pessoa7_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa8_idade= factor(banco_dados_jna_2025$pessoa8_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa9_idade= factor(banco_dados_jna_2025$pessoa9_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa10_idade= factor(banco_dados_jna_2025$pessoa10_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa11_idade= factor(banco_dados_jna_2025$pessoa11_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)
# banco_dados_jna_2025$pessoa12_idade= factor(banco_dados_jna_2025$pessoa12_idade,
#                                            levels = v40$levels,
#                                            labels = v40$labels)

#pessoas idade
banco_dados_jna_2025$pessoa1_idade_anos  = factor(banco_dados_jna_2025$pessoa1_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa2_idade_anos  = factor(banco_dados_jna_2025$pessoa2_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa3_idade_anos  = factor(banco_dados_jna_2025$pessoa3_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa4_idade_anos  = factor(banco_dados_jna_2025$pessoa4_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa5_idade_anos  = factor(banco_dados_jna_2025$pessoa5_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa6_idade_anos  = factor(banco_dados_jna_2025$pessoa6_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa7_idade_anos  = factor(banco_dados_jna_2025$pessoa7_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa8_idade_anos  = factor(banco_dados_jna_2025$pessoa8_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa9_idade_anos  = factor(banco_dados_jna_2025$pessoa9_idade_anos,  levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa10_idade_anos = factor(banco_dados_jna_2025$pessoa10_idade_anos, levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa11_idade_anos = factor(banco_dados_jna_2025$pessoa11_idade_anos, levels = v40$levels, labels = v40$labels)
banco_dados_jna_2025$pessoa12_idade_anos = factor(banco_dados_jna_2025$pessoa12_idade_anos, levels = v40$levels, labels = v40$labels)


#pessoas genero

banco_dados_jna_2025$pessoa1_genero= factor(banco_dados_jna_2025$pessoa1_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa2_genero= factor(banco_dados_jna_2025$pessoa2_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa3_genero= factor(banco_dados_jna_2025$pessoa3_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa4_genero= factor(banco_dados_jna_2025$pessoa4_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa5_genero= factor(banco_dados_jna_2025$pessoa5_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa6_genero= factor(banco_dados_jna_2025$pessoa6_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa7_genero= factor(banco_dados_jna_2025$pessoa7_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa8_genero= factor(banco_dados_jna_2025$pessoa8_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa9_genero= factor(banco_dados_jna_2025$pessoa9_genero,
                                           levels = v7_37$levels,
                                           labels = v7_37$labels)
banco_dados_jna_2025$pessoa10_genero= factor(banco_dados_jna_2025$pessoa10_genero,
                                            levels = v7_37$levels,
                                            labels = v7_37$labels)
banco_dados_jna_2025$pessoa11_genero= factor(banco_dados_jna_2025$pessoa11_genero,
                                            levels = v7_37$levels,
                                            labels = v7_37$labels)
banco_dados_jna_2025$pessoa12_genero= factor(banco_dados_jna_2025$pessoa12_genero,
                                            levels = v7_37$levels,
                                            labels = v7_37$labels)

#pessoas parentesco

banco_dados_jna_2025$pessoa1_parentesco= factor(banco_dados_jna_2025$pessoa1_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa2_parentesco= factor(banco_dados_jna_2025$pessoa2_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa3_parentesco= factor(banco_dados_jna_2025$pessoa3_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa4_parentesco= factor(banco_dados_jna_2025$pessoa4_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa5_parentesco= factor(banco_dados_jna_2025$pessoa5_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa6_parentesco= factor(banco_dados_jna_2025$pessoa6_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa7_parentesco= factor(banco_dados_jna_2025$pessoa7_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa8_parentesco= factor(banco_dados_jna_2025$pessoa8_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa9_parentesco= factor(banco_dados_jna_2025$pessoa9_parentesco,
                                            levels = v14_39$levels,
                                            labels = v14_39$labels)
banco_dados_jna_2025$pessoa10_parentesco= factor(banco_dados_jna_2025$pessoa10_parentesco,
                                             levels = v14_39$levels,
                                             labels = v14_39$labels)
banco_dados_jna_2025$pessoa11_parentesco= factor(banco_dados_jna_2025$pessoa11_parentesco,
                                             levels = v14_39$levels,
                                             labels = v14_39$labels)
banco_dados_jna_2025$pessoa12_parentesco= factor(banco_dados_jna_2025$pessoa12_parentesco,
                                             levels = v14_39$levels,
                                             labels = v14_39$labels)

#pessoas nacionalidade

banco_dados_jna_2025$pessoa1_nacionalidade= factor(banco_dados_jna_2025$pessoa1_nacionalidade,
                                                     levels = v52_1$levels,
                                                     labels = v52_1$labels)
banco_dados_jna_2025$pessoa2_nacionalidade= factor(banco_dados_jna_2025$pessoa2_nacionalidade,
                                                     levels = v52_2$levels,
                                                     labels = v52_2$labels)
banco_dados_jna_2025$pessoa3_nacionalidade= factor(banco_dados_jna_2025$pessoa3_nacionalidade,
                                                     levels = v52_3$levels,
                                                     labels = v52_3$labels)
banco_dados_jna_2025$pessoa4_nacionalidade= factor(banco_dados_jna_2025$pessoa4_nacionalidade,
                                                     levels = v52_4$levels,
                                                     labels = v52_4$labels)
banco_dados_jna_2025$pessoa5_nacionalidade= factor(banco_dados_jna_2025$pessoa5_nacionalidade,
                                                     levels = v52_5$levels,
                                                     labels = v52_5$labels)
banco_dados_jna_2025$pessoa6_nacionalidade= factor(banco_dados_jna_2025$pessoa6_nacionalidade,
                                                     levels = v52_6$levels,
                                                     labels =v52_6$labels)
banco_dados_jna_2025$pessoa7_nacionalidade= factor(banco_dados_jna_2025$pessoa7_nacionalidade,
                                                     levels = v52_7$levels,
                                                     labels = v52_7$labels)
banco_dados_jna_2025$pessoa8_nacionalidade= factor(banco_dados_jna_2025$pessoa8_nacionalidade,
                                                     levels = v52_8$levels,
                                                     labels = v52_8$labels)
banco_dados_jna_2025$pessoa9_nacionalidade= factor(banco_dados_jna_2025$pessoa9_nacionalidade,
                                                     levels = v52_9$levels,
                                                     labels = v52_9$labels)
banco_dados_jna_2025$pessoa10_nacionalidade= factor(banco_dados_jna_2025$pessoa10_nacionalidade,
                                                      levels = v52_10$levels,
                                                      labels = v52_10$labels)
banco_dados_jna_2025$pessoa11_nacionalidade= factor(banco_dados_jna_2025$pessoa11_nacionalidade,
                                                      levels = v52_11$levels,
                                                      labels = v52_11$labels)
banco_dados_jna_2025$pessoa12_nacionalidade= factor(banco_dados_jna_2025$pessoa12_nacionalidade,
                                                      levels = v52_12$levels,
                                                      labels = v52_12$labels)


#pessoas gravidas lactantes

banco_dados_jna_2025$pessoa1_gravida_lactante= factor(banco_dados_jna_2025$pessoa1_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa2_gravida_lactante= factor(banco_dados_jna_2025$pessoa2_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa3_gravida_lactante= factor(banco_dados_jna_2025$pessoa3_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa4_gravida_lactante= factor(banco_dados_jna_2025$pessoa4_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa5_gravida_lactante= factor(banco_dados_jna_2025$pessoa5_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa6_gravida_lactante= factor(banco_dados_jna_2025$pessoa6_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels =v16_41$labels)
banco_dados_jna_2025$pessoa7_gravida_lactante= factor(banco_dados_jna_2025$pessoa7_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa8_gravida_lactante= factor(banco_dados_jna_2025$pessoa8_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa9_gravida_lactante= factor(banco_dados_jna_2025$pessoa9_gravida_lactante,
                                           levels = v16_41$levels,
                                           labels = v16_41$labels)
banco_dados_jna_2025$pessoa10_gravida_lactante= factor(banco_dados_jna_2025$pessoa10_gravida_lactante,
                                            levels = v16_41$levels,
                                            labels = v16_41$labels)
banco_dados_jna_2025$pessoa11_gravida_lactante= factor(banco_dados_jna_2025$pessoa11_gravida_lactante,
                                            levels = v16_41$levels,
                                            labels = v16_41$labels)
banco_dados_jna_2025$pessoa12_gravida_lactante= factor(banco_dados_jna_2025$pessoa12_gravida_lactante,
                                            levels = v16_41$levels,
                                            labels = v16_41$labels)


#pessoas gravida servico nutricao

banco_dados_jna_2025$pessoa1_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa1_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa2_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa2_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa3_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa3_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa4_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa4_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa5_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa5_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa6_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa6_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels =v17_42_43$labels)
banco_dados_jna_2025$pessoa7_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa7_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa8_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa8_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa9_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa9_gravida_servico_nutricao,
                                                      levels = v17_42_43$levels,
                                                      labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa10_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa10_gravida_servico_nutricao,
                                                       levels = v17_42_43$levels,
                                                       labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa11_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa11_gravida_servico_nutricao,
                                                       levels = v17_42_43$levels,
                                                       labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa12_gravida_servico_nutricao= factor(banco_dados_jna_2025$pessoa12_gravida_servico_nutricao,
                                                       levels = v17_42_43$levels,
                                                       labels = v17_42_43$labels)

#pessoas servico nutricao criancas

banco_dados_jna_2025$pessoa1_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa1_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa2_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa2_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa3_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa3_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa4_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa4_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa5_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa5_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa6_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa6_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels =v17_42_43$labels)
banco_dados_jna_2025$pessoa7_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa7_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa8_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa8_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa9_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa9_servico_nutricao_criancas,
                                                              levels = v17_42_43$levels,
                                                              labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa10_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa10_servico_nutricao_criancas,
                                                               levels = v17_42_43$levels,
                                                               labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa11_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa11_servico_nutricao_criancas,
                                                               levels = v17_42_43$levels,
                                                               labels = v17_42_43$labels)
banco_dados_jna_2025$pessoa12_servico_nutricao_criancas= factor(banco_dados_jna_2025$pessoa12_servico_nutricao_criancas,
                                                               levels = v17_42_43$levels,
                                                               labels = v17_42_43$labels)

#pessoas etnia

banco_dados_jna_2025$pessoa1_etnia= factor(banco_dados_jna_2025$pessoa1_etnia,
                                                              levels = v44_1$levels,
                                                              labels = v44_1$labels)
banco_dados_jna_2025$pessoa2_etnia= factor(banco_dados_jna_2025$pessoa2_etnia,
                                                              levels = v44_2$levels,
                                                              labels = v44_2$labels)
banco_dados_jna_2025$pessoa3_etnia= factor(banco_dados_jna_2025$pessoa3_etnia,
                                                              levels = v44_3$levels,
                                                              labels = v44_3$labels)
banco_dados_jna_2025$pessoa4_etnia= factor(banco_dados_jna_2025$pessoa4_etnia,
                                                              levels = v44_4$levels,
                                                              labels = v44_4$labels)
banco_dados_jna_2025$pessoa5_etnia= factor(banco_dados_jna_2025$pessoa5_etnia,
                                                              levels = v44_5$levels,
                                                              labels = v44_5$labels)
banco_dados_jna_2025$pessoa6_etnia= factor(banco_dados_jna_2025$pessoa6_etnia,
                                                              levels = v44_6$levels,
                                                              labels =v44_6$labels)
banco_dados_jna_2025$pessoa7_etnia= factor(banco_dados_jna_2025$pessoa7_etnia,
                                                              levels = v44_7$levels,
                                                              labels = v44_7$labels)
banco_dados_jna_2025$pessoa8_etnia= factor(banco_dados_jna_2025$pessoa8_etnia,
                                                              levels = v44_8$levels,
                                                              labels = v44_8$labels)
banco_dados_jna_2025$pessoa9_etnia= factor(banco_dados_jna_2025$pessoa9_etnia,
                                                              levels = v44_9$levels,
                                                              labels = v44_9$labels)
banco_dados_jna_2025$pessoa10_etnia= factor(banco_dados_jna_2025$pessoa10_etnia,
                                                               levels = v44_10$levels,
                                                               labels = v44_10$labels)
banco_dados_jna_2025$pessoa11_etnia= factor(banco_dados_jna_2025$pessoa11_etnia,
                                                               levels = v44_11$levels,
                                                               labels = v44_11$labels)
banco_dados_jna_2025$pessoa12_etnia= factor(banco_dados_jna_2025$pessoa12_etnia,
                                                               levels = v44_12$levels,
                                                               labels = v44_12$labels)

#pessoas deficiencia

banco_dados_jna_2025$pessoa1_deficiencia= factor(banco_dados_jna_2025$pessoa1_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa2_deficiencia= factor(banco_dados_jna_2025$pessoa2_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa3_deficiencia= factor(banco_dados_jna_2025$pessoa3_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa4_deficiencia= factor(banco_dados_jna_2025$pessoa4_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa5_deficiencia= factor(banco_dados_jna_2025$pessoa5_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa6_deficiencia= factor(banco_dados_jna_2025$pessoa6_deficiencia,
                                           levels = v19_45$levels,
                                           labels =v19_45$labels)
banco_dados_jna_2025$pessoa7_deficiencia= factor(banco_dados_jna_2025$pessoa7_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa8_deficiencia= factor(banco_dados_jna_2025$pessoa8_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa9_deficiencia= factor(banco_dados_jna_2025$pessoa9_deficiencia,
                                           levels = v19_45$levels,
                                           labels = v19_45$labels)
banco_dados_jna_2025$pessoa10_deficiencia= factor(banco_dados_jna_2025$pessoa10_deficiencia,
                                            levels = v19_45$levels,
                                            labels = v19_45$labels)
banco_dados_jna_2025$pessoa11_deficiencia= factor(banco_dados_jna_2025$pessoa11_deficiencia,
                                            levels = v19_45$levels,
                                            labels = v19_45$labels)
banco_dados_jna_2025$pessoa12_deficiencia= factor(banco_dados_jna_2025$pessoa12_deficiencia,
                                            levels = v19_45$levels,
                                            labels = v19_45$labels)


#pessoas deficiencia fisica

banco_dados_jna_2025$pessoa1_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa1_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa2_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa2_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa3_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa3_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa4_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa4_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa5_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa5_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa6_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa6_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels =v20_46$labels)
banco_dados_jna_2025$pessoa7_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa7_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa8_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa8_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa9_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa9_deficiencia_fisica,
                                                 levels = v20_46$levels,
                                                 labels = v20_46$labels)
banco_dados_jna_2025$pessoa10_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa10_deficiencia_fisica,
                                                  levels = v20_46$levels,
                                                  labels = v20_46$labels)
banco_dados_jna_2025$pessoa11_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa11_deficiencia_fisica,
                                                  levels = v20_46$levels,
                                                  labels = v20_46$labels)
banco_dados_jna_2025$pessoa12_deficiencia_fisica= factor(banco_dados_jna_2025$pessoa12_deficiencia_fisica,
                                                  levels = v20_46$levels,
                                                  labels = v20_46$labels)


#pessoas deficiencia mental

banco_dados_jna_2025$pessoa1_deficiencia_mental= factor(banco_dados_jna_2025$pessoa1_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa2_deficiencia_mental= factor(banco_dados_jna_2025$pessoa2_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa3_deficiencia_mental= factor(banco_dados_jna_2025$pessoa3_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa4_deficiencia_mental= factor(banco_dados_jna_2025$pessoa4_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa5_deficiencia_mental= factor(banco_dados_jna_2025$pessoa5_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa6_deficiencia_mental= factor(banco_dados_jna_2025$pessoa6_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels =v21_47$labels)
banco_dados_jna_2025$pessoa7_deficiencia_mental= factor(banco_dados_jna_2025$pessoa7_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa8_deficiencia_mental= factor(banco_dados_jna_2025$pessoa8_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa9_deficiencia_mental= factor(banco_dados_jna_2025$pessoa9_deficiencia_mental,
                                                        levels = v21_47$levels,
                                                        labels = v21_47$labels)
banco_dados_jna_2025$pessoa10_deficiencia_mental= factor(banco_dados_jna_2025$pessoa10_deficiencia_mental,
                                                         levels = v21_47$levels,
                                                         labels = v21_47$labels)
banco_dados_jna_2025$pessoa11_deficiencia_mental= factor(banco_dados_jna_2025$pessoa11_deficiencia_mental,
                                                         levels = v21_47$levels,
                                                         labels = v21_47$labels)
banco_dados_jna_2025$pessoa12_deficiencia_mental= factor(banco_dados_jna_2025$pessoa12_deficiencia_mental,
                                                         levels = v21_47$levels,
                                                         labels = v21_47$labels)

#pessoas deficiencia auditiva

banco_dados_jna_2025$pessoa1_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa1_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa2_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa2_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa3_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa3_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa4_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa4_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa5_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa5_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa6_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa6_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels =v22_48$labels)
banco_dados_jna_2025$pessoa7_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa7_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa8_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa8_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa9_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa9_deficiencia_auditiva,
                                                        levels = v22_48$levels,
                                                        labels = v22_48$labels)
banco_dados_jna_2025$pessoa10_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa10_deficiencia_auditiva,
                                                         levels = v22_48$levels,
                                                         labels = v22_48$labels)
banco_dados_jna_2025$pessoa11_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa11_deficiencia_auditiva,
                                                         levels = v22_48$levels,
                                                         labels = v22_48$labels)
banco_dados_jna_2025$pessoa12_deficiencia_auditiva= factor(banco_dados_jna_2025$pessoa12_deficiencia_auditiva,
                                                         levels = v22_48$levels,
                                                         labels = v22_48$labels)


#pessoas deficiencia visual

banco_dados_jna_2025$pessoa1_deficiencia_visual= factor(banco_dados_jna_2025$pessoa1_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa2_deficiencia_visual= factor(banco_dados_jna_2025$pessoa2_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa3_deficiencia_visual= factor(banco_dados_jna_2025$pessoa3_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa4_deficiencia_visual= factor(banco_dados_jna_2025$pessoa4_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa5_deficiencia_visual= factor(banco_dados_jna_2025$pessoa5_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa6_deficiencia_visual= factor(banco_dados_jna_2025$pessoa6_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels =v23_49$labels)
banco_dados_jna_2025$pessoa7_deficiencia_visual= factor(banco_dados_jna_2025$pessoa7_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa8_deficiencia_visual= factor(banco_dados_jna_2025$pessoa8_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa9_deficiencia_visual= factor(banco_dados_jna_2025$pessoa9_deficiencia_visual,
                                                          levels = v23_49$levels,
                                                          labels = v23_49$labels)
banco_dados_jna_2025$pessoa10_deficiencia_visual= factor(banco_dados_jna_2025$pessoa10_deficiencia_visual,
                                                           levels = v23_49$levels,
                                                           labels = v23_49$labels)
banco_dados_jna_2025$pessoa11_deficiencia_visual= factor(banco_dados_jna_2025$pessoa11_deficiencia_visual,
                                                           levels = v23_49$levels,
                                                           labels = v23_49$labels)
banco_dados_jna_2025$pessoa12_deficiencia_visual= factor(banco_dados_jna_2025$pessoa12_deficiencia_visual,
                                                           levels = v23_49$levels,
                                                           labels = v23_49$labels)

#pessoas deficiencia autismo

banco_dados_jna_2025$pessoa1_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa1_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa2_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa2_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa3_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa3_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa4_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa4_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa5_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa5_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa6_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa6_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels =v24_50$labels)
banco_dados_jna_2025$pessoa7_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa7_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa8_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa8_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa9_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa9_deficiencia_autismo,
                                                        levels = v24_50$levels,
                                                        labels = v24_50$labels)
banco_dados_jna_2025$pessoa10_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa10_deficiencia_autismo,
                                                         levels = v24_50$levels,
                                                         labels = v24_50$labels)
banco_dados_jna_2025$pessoa11_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa11_deficiencia_autismo,
                                                         levels = v24_50$levels,
                                                         labels = v24_50$labels)
banco_dados_jna_2025$pessoa12_deficiencia_autismo= factor(banco_dados_jna_2025$pessoa12_deficiencia_autismo,
                                                         levels = v24_50$levels,
                                                         labels = v24_50$labels)



#pessoas deficiencia outro

banco_dados_jna_2025$pessoa1_deficiencia_outro= factor(banco_dados_jna_2025$pessoa1_deficiencia_outro,
                                                         levels = v51_1$levels,
                                                         labels = v51_1$labels)
banco_dados_jna_2025$pessoa2_deficiencia_outro= factor(banco_dados_jna_2025$pessoa2_deficiencia_outro,
                                                         levels = v51_2$levels,
                                                         labels = v51_2$labels)
banco_dados_jna_2025$pessoa3_deficiencia_outro= factor(banco_dados_jna_2025$pessoa3_deficiencia_outro,
                                                         levels = v51_3$levels,
                                                         labels = v51_3$labels)
banco_dados_jna_2025$pessoa4_deficiencia_outro= factor(banco_dados_jna_2025$pessoa4_deficiencia_outro,
                                                         levels = v51_4$levels,
                                                          labels = v51_4$labels)

#pessoas documento residencia e refugio

banco_dados_jna_2025$pessoa1_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa1_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa2_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa2_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa3_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa3_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa4_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa4_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa5_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa5_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa6_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa6_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels =v26_53$labels)
banco_dados_jna_2025$pessoa7_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa7_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa8_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa8_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa9_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa9_documento_refugio_residencia,
                                                       levels = v26_53$levels,
                                                       labels = v26_53$labels)
banco_dados_jna_2025$pessoa10_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa10_documento_refugio_residencia,
                                                        levels = v26_53$levels,
                                                        labels = v26_53$labels)
banco_dados_jna_2025$pessoa11_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa11_documento_refugio_residencia,
                                                        levels = v26_53$levels,
                                                        labels = v26_53$labels)
banco_dados_jna_2025$pessoa12_documento_refugio_residencia= factor(banco_dados_jna_2025$pessoa12_documento_refugio_residencia,
                                                        levels = v26_53$levels,
                                                        labels = v26_53$labels)


#pessoas cpf

banco_dados_jna_2025$pessoa1_cpf= factor(banco_dados_jna_2025$pessoa1_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa2_cpf= factor(banco_dados_jna_2025$pessoa2_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa3_cpf= factor(banco_dados_jna_2025$pessoa3_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa4_cpf= factor(banco_dados_jna_2025$pessoa4_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa5_cpf= factor(banco_dados_jna_2025$pessoa5_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa6_cpf= factor(banco_dados_jna_2025$pessoa6_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels =v27_54$labels)
banco_dados_jna_2025$pessoa7_cpf= factor(banco_dados_jna_2025$pessoa7_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa8_cpf= factor(banco_dados_jna_2025$pessoa8_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa9_cpf= factor(banco_dados_jna_2025$pessoa9_cpf,
                                                                  levels = v27_54$levels,
                                                                  labels = v27_54$labels)
banco_dados_jna_2025$pessoa10_cpf= factor(banco_dados_jna_2025$pessoa10_cpf,
                                                                   levels = v27_54$levels,
                                                                   labels = v27_54$labels)
banco_dados_jna_2025$pessoa11_cpf= factor(banco_dados_jna_2025$pessoa11_cpf,
                                                                   levels = v27_54$levels,
                                                                   labels = v27_54$labels)
banco_dados_jna_2025$pessoa12_cpf= factor(banco_dados_jna_2025$pessoa12_cpf,
                                                                   levels = v27_54$levels,
                                                                   labels = v27_54$labels)

#pessoas cnh

banco_dados_jna_2025$pessoa1_cnh= factor(banco_dados_jna_2025$pessoa1_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa2_cnh= factor(banco_dados_jna_2025$pessoa2_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa3_cnh= factor(banco_dados_jna_2025$pessoa3_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa4_cnh= factor(banco_dados_jna_2025$pessoa4_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa5_cnh= factor(banco_dados_jna_2025$pessoa5_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa6_cnh= factor(banco_dados_jna_2025$pessoa6_cnh,
                                         levels = v28_55$levels,
                                         labels =v28_55$labels)
banco_dados_jna_2025$pessoa7_cnh= factor(banco_dados_jna_2025$pessoa7_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa8_cnh= factor(banco_dados_jna_2025$pessoa8_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa9_cnh= factor(banco_dados_jna_2025$pessoa9_cnh,
                                         levels = v28_55$levels,
                                         labels = v28_55$labels)
banco_dados_jna_2025$pessoa10_cnh= factor(banco_dados_jna_2025$pessoa10_cnh,
                                          levels = v28_55$levels,
                                          labels = v28_55$labels)
banco_dados_jna_2025$pessoa11_cnh= factor(banco_dados_jna_2025$pessoa11_cnh,
                                          levels = v28_55$levels,
                                          labels = v28_55$labels)
banco_dados_jna_2025$pessoa12_cnh= factor(banco_dados_jna_2025$pessoa12_cnh,
                                          levels = v28_55$levels,
                                          labels = v28_55$labels)

#pessoas certidao nascimento

banco_dados_jna_2025$pessoa1_certidao_nascimento= factor(banco_dados_jna_2025$pessoa1_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa2_certidao_nascimento= factor(banco_dados_jna_2025$pessoa2_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa3_certidao_nascimento= factor(banco_dados_jna_2025$pessoa3_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa4_certidao_nascimento= factor(banco_dados_jna_2025$pessoa4_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa5_certidao_nascimento= factor(banco_dados_jna_2025$pessoa5_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa6_certidao_nascimento= factor(banco_dados_jna_2025$pessoa6_certidao_nascimento,
                                         levels = v56$levels,
                                         labels =v56$labels)
banco_dados_jna_2025$pessoa7_certidao_nascimento= factor(banco_dados_jna_2025$pessoa7_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa8_certidao_nascimento= factor(banco_dados_jna_2025$pessoa8_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa9_certidao_nascimento= factor(banco_dados_jna_2025$pessoa9_certidao_nascimento,
                                         levels = v56$levels,
                                         labels = v56$labels)
banco_dados_jna_2025$pessoa10_certidao_nascimento= factor(banco_dados_jna_2025$pessoa10_certidao_nascimento,
                                          levels = v56$levels,
                                          labels = v56$labels)
banco_dados_jna_2025$pessoa11_certidao_nascimento= factor(banco_dados_jna_2025$pessoa11_certidao_nascimento,
                                          levels = v56$levels,
                                          labels = v56$labels)
banco_dados_jna_2025$pessoa12_certidao_nascimento= factor(banco_dados_jna_2025$pessoa12_certidao_nascimento,
                                          levels = v56$levels,
                                          labels = v56$labels)

#pessoas sus

banco_dados_jna_2025$pessoa1_sus= factor(banco_dados_jna_2025$pessoa1_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa2_sus= factor(banco_dados_jna_2025$pessoa2_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa3_sus= factor(banco_dados_jna_2025$pessoa3_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa4_sus= factor(banco_dados_jna_2025$pessoa4_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa5_sus= factor(banco_dados_jna_2025$pessoa5_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa6_sus= factor(banco_dados_jna_2025$pessoa6_sus,
                                         levels = v29_57$levels,
                                         labels =v29_57$labels)
banco_dados_jna_2025$pessoa7_sus= factor(banco_dados_jna_2025$pessoa7_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa8_sus= factor(banco_dados_jna_2025$pessoa8_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa9_sus= factor(banco_dados_jna_2025$pessoa9_sus,
                                         levels = v29_57$levels,
                                         labels = v29_57$labels)
banco_dados_jna_2025$pessoa10_sus= factor(banco_dados_jna_2025$pessoa10_sus,
                                          levels = v29_57$levels,
                                          labels = v29_57$labels)
banco_dados_jna_2025$pessoa11_sus= factor(banco_dados_jna_2025$pessoa11_sus,
                                          levels = v29_57$levels,
                                          labels = v29_57$labels)
banco_dados_jna_2025$pessoa12_sus= factor(banco_dados_jna_2025$pessoa12_sus,
                                          levels = v29_57$levels,
                                          labels = v29_57$labels)

#pessoas escola

banco_dados_jna_2025$pessoa1_escola= factor(banco_dados_jna_2025$pessoa1_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa2_escola= factor(banco_dados_jna_2025$pessoa2_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa3_escola= factor(banco_dados_jna_2025$pessoa3_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa4_escola= factor(banco_dados_jna_2025$pessoa4_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa5_escola= factor(banco_dados_jna_2025$pessoa5_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa6_escola= factor(banco_dados_jna_2025$pessoa6_escola,
                                         levels = v30_58$levels,
                                         labels =v30_58$labels)
banco_dados_jna_2025$pessoa7_escola= factor(banco_dados_jna_2025$pessoa7_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa8_escola= factor(banco_dados_jna_2025$pessoa8_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa9_escola= factor(banco_dados_jna_2025$pessoa9_escola,
                                         levels = v30_58$levels,
                                         labels = v30_58$labels)
banco_dados_jna_2025$pessoa10_escola= factor(banco_dados_jna_2025$pessoa10_escola,
                                          levels = v30_58$levels,
                                          labels = v30_58$labels)
banco_dados_jna_2025$pessoa11_escola= factor(banco_dados_jna_2025$pessoa11_escola,
                                          levels = v30_58$levels,
                                          labels = v30_58$labels)
banco_dados_jna_2025$pessoa12_escola= factor(banco_dados_jna_2025$pessoa12_escola,
                                          levels = v30_58$levels,
                                          labels = v30_58$labels)

#pessoas escola barreira

banco_dados_jna_2025$pessoa1_barreira_escola= factor(banco_dados_jna_2025$pessoa1_barreira_escola,
                                            levels = v59_1$levels,
                                            labels = v59_1$labels)
banco_dados_jna_2025$pessoa2_barreira_escola= factor(banco_dados_jna_2025$pessoa2_barreira_escola,
                                            levels = v59_2$levels,
                                            labels = v59_2$labels)
banco_dados_jna_2025$pessoa3_barreira_escola= factor(banco_dados_jna_2025$pessoa3_barreira_escola,
                                            levels = v59_3$levels,
                                            labels = v59_3$labels)
banco_dados_jna_2025$pessoa4_barreira_escola= factor(banco_dados_jna_2025$pessoa4_barreira_escola,
                                            levels = v59_4$levels,
                                            labels = v59_4$labels)
banco_dados_jna_2025$pessoa5_barreira_escola= factor(banco_dados_jna_2025$pessoa5_barreira_escola,
                                            levels = v59_5$levels,
                                            labels = v59_5$labels)
banco_dados_jna_2025$pessoa6_barreira_escola= factor(banco_dados_jna_2025$pessoa6_barreira_escola,
                                            levels = v59_6$levels,
                                            labels =v59_6$labels)
banco_dados_jna_2025$pessoa7_barreira_escola= factor(banco_dados_jna_2025$pessoa7_barreira_escola,
                                            levels = v59_7$levels,
                                            labels = v59_7$labels)
banco_dados_jna_2025$pessoa8_barreira_escola= factor(banco_dados_jna_2025$pessoa8_barreira_escola,
                                            levels = v59_8$levels,
                                            labels = v59_8$labels)
banco_dados_jna_2025$pessoa9_barreira_escola= factor(banco_dados_jna_2025$pessoa9_barreira_escola,
                                            levels = v59_9$levels,
                                            labels = v59_9$labels)
banco_dados_jna_2025$pessoa10_barreira_escola= factor(banco_dados_jna_2025$pessoa10_barreira_escola,
                                             levels = v59_10$levels,
                                             labels = v59_10$labels)
banco_dados_jna_2025$pessoa11_barreira_escola= factor(banco_dados_jna_2025$pessoa11_barreira_escola,
                                             levels = v59_11$levels,
                                             labels = v59_11$labels)
banco_dados_jna_2025$pessoa12_barreira_escola= factor(banco_dados_jna_2025$pessoa12_barreira_escola,
                                             levels = v59_12$levels,
                                             labels = v59_12$labels)

#pessoas grau escolar

banco_dados_jna_2025$pessoa1_grau_escolar= factor(banco_dados_jna_2025$pessoa1_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa2_grau_escolar= factor(banco_dados_jna_2025$pessoa2_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa3_grau_escolar= factor(banco_dados_jna_2025$pessoa3_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa4_grau_escolar= factor(banco_dados_jna_2025$pessoa4_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa5_grau_escolar= factor(banco_dados_jna_2025$pessoa5_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa6_grau_escolar= factor(banco_dados_jna_2025$pessoa6_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels =v32_60$labels)
banco_dados_jna_2025$pessoa7_grau_escolar= factor(banco_dados_jna_2025$pessoa7_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa8_grau_escolar= factor(banco_dados_jna_2025$pessoa8_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa9_grau_escolar= factor(banco_dados_jna_2025$pessoa9_grau_escolar,
                                                           levels = v32_60$levels,
                                                           labels = v32_60$labels)
banco_dados_jna_2025$pessoa10_grau_escolar= factor(banco_dados_jna_2025$pessoa10_grau_escolar,
                                                            levels = v32_60$levels,
                                                            labels = v32_60$labels)
banco_dados_jna_2025$pessoa11_grau_escolar= factor(banco_dados_jna_2025$pessoa11_grau_escolar,
                                                            levels = v32_60$levels,
                                                            labels = v32_60$labels)
banco_dados_jna_2025$pessoa12_grau_escolar= factor(banco_dados_jna_2025$pessoa12_grau_escolar,
                                                            levels = v32_60$levels,
                                                            labels = v32_60$labels)



#pessoas grau escolar completo

banco_dados_jna_2025$pessoa1_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa1_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa2_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa2_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa3_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa3_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa4_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa4_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa5_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa5_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa6_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa6_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels =v33_61$labels)
banco_dados_jna_2025$pessoa7_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa7_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa8_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa8_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa9_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa9_grau_escolar_completo,
                                                     levels = v33_61$levels,
                                                     labels = v33_61$labels)
banco_dados_jna_2025$pessoa10_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa10_grau_escolar_completo,
                                                      levels = v33_61$levels,
                                                      labels = v33_61$labels)
banco_dados_jna_2025$pessoa11_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa11_grau_escolar_completo,
                                                      levels = v33_61$levels,
                                                      labels = v33_61$labels)
banco_dados_jna_2025$pessoa12_grau_escolar_completo= factor(banco_dados_jna_2025$pessoa12_grau_escolar_completo,
                                                      levels = v33_61$levels,
                                                      labels = v33_61$labels)


#pessoas trabalhando

banco_dados_jna_2025$pessoa1_trabalhando= factor(banco_dados_jna_2025$pessoa1_trabalhando,
                                                           levels = v62_1$levels,
                                                           labels = v62_1$labels)
banco_dados_jna_2025$pessoa2_trabalhando= factor(banco_dados_jna_2025$pessoa2_trabalhando,
                                                           levels = v62_2$levels,
                                                           labels = v62_2$labels)
banco_dados_jna_2025$pessoa3_trabalhando= factor(banco_dados_jna_2025$pessoa3_trabalhando,
                                                           levels = v62_3$levels,
                                                           labels = v62_3$labels)
banco_dados_jna_2025$pessoa4_trabalhando= factor(banco_dados_jna_2025$pessoa4_trabalhando,
                                                           levels = v62_4$levels,
                                                           labels = v62_4$labels)
banco_dados_jna_2025$pessoa5_trabalhando= factor(banco_dados_jna_2025$pessoa5_trabalhando,
                                                           levels = v62_5$levels,
                                                           labels = v62_5$labels)
banco_dados_jna_2025$pessoa6_trabalhando= factor(banco_dados_jna_2025$pessoa6_trabalhando,
                                                           levels = v62_6$levels,
                                                           labels =v62_6$labels)
banco_dados_jna_2025$pessoa7_trabalhando= factor(banco_dados_jna_2025$pessoa7_trabalhando,
                                                           levels = v62_7$levels,
                                                           labels = v62_7$labels)
banco_dados_jna_2025$pessoa8_trabalhando= factor(banco_dados_jna_2025$pessoa8_trabalhando,
                                                           levels = v62_8$levels,
                                                           labels = v62_8$labels)
banco_dados_jna_2025$pessoa9_trabalhando= factor(banco_dados_jna_2025$pessoa9_trabalhando,
                                                           levels = v62_9$levels,
                                                           labels = v62_9$labels)
banco_dados_jna_2025$pessoa10_trabalhando= factor(banco_dados_jna_2025$pessoa10_trabalhando,
                                                            levels = v62_10$levels,
                                                            labels = v62_10$labels)
banco_dados_jna_2025$pessoa11_trabalhando= factor(banco_dados_jna_2025$pessoa11_trabalhando,
                                                            levels = v62_11$levels,
                                                            labels = v62_11$labels)
banco_dados_jna_2025$pessoa12_trabalhando= factor(banco_dados_jna_2025$pessoa12_trabalhando,
                                                            levels = v62_12$levels,
                                                            labels = v62_12$labels)


#pessoas trabalhando contrato

banco_dados_jna_2025$pessoa1_contrato_trabalho= factor(banco_dados_jna_2025$pessoa1_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa2_contrato_trabalho= factor(banco_dados_jna_2025$pessoa2_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa3_contrato_trabalho= factor(banco_dados_jna_2025$pessoa3_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa4_contrato_trabalho= factor(banco_dados_jna_2025$pessoa4_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa5_contrato_trabalho= factor(banco_dados_jna_2025$pessoa5_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa6_contrato_trabalho= factor(banco_dados_jna_2025$pessoa6_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels =v35_63$labels)
banco_dados_jna_2025$pessoa7_contrato_trabalho= factor(banco_dados_jna_2025$pessoa7_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa8_contrato_trabalho= factor(banco_dados_jna_2025$pessoa8_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa9_contrato_trabalho= factor(banco_dados_jna_2025$pessoa9_contrato_trabalho,
                                                 levels = v35_63$levels,
                                                 labels = v35_63$labels)
banco_dados_jna_2025$pessoa10_contrato_trabalho= factor(banco_dados_jna_2025$pessoa10_contrato_trabalho,
                                                  levels = v35_63$levels,
                                                  labels = v35_63$labels)
banco_dados_jna_2025$pessoa11_contrato_trabalho= factor(banco_dados_jna_2025$pessoa11_contrato_trabalho,
                                                  levels = v35_63$levels,
                                                  labels = v35_63$labels)
banco_dados_jna_2025$pessoa12_contrato_trabalho= factor(banco_dados_jna_2025$pessoa12_contrato_trabalho,
                                                  levels = v35_63$levels,
                                                  labels = v35_63$labels)

#pessoas microempresa

banco_dados_jna_2025$pessoa1_microempresa= factor(banco_dados_jna_2025$pessoa1_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa2_microempresa= factor(banco_dados_jna_2025$pessoa2_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa3_microempresa= factor(banco_dados_jna_2025$pessoa3_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa4_microempresa= factor(banco_dados_jna_2025$pessoa4_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa5_microempresa= factor(banco_dados_jna_2025$pessoa5_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa6_microempresa= factor(banco_dados_jna_2025$pessoa6_microempresa,
                                                       levels = v36_64$levels,
                                                       labels =v36_64$labels)
banco_dados_jna_2025$pessoa7_microempresa= factor(banco_dados_jna_2025$pessoa7_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa8_microempresa= factor(banco_dados_jna_2025$pessoa8_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa9_microempresa= factor(banco_dados_jna_2025$pessoa9_microempresa,
                                                       levels = v36_64$levels,
                                                       labels = v36_64$labels)
banco_dados_jna_2025$pessoa10_microempresa= factor(banco_dados_jna_2025$pessoa10_microempresa,
                                                        levels = v36_64$levels,
                                                        labels = v36_64$labels)
banco_dados_jna_2025$pessoa11_microempresa= factor(banco_dados_jna_2025$pessoa11_microempresa,
                                                        levels = v36_64$levels,
                                                        labels = v36_64$labels)
banco_dados_jna_2025$pessoa12_microempresa= factor(banco_dados_jna_2025$pessoa12_microempresa,
                                                        levels = v36_64$levels,
                                                        labels = v36_64$labels)

# menores 18 anos
banco_dados_jna_2025$menores_18_hogar= factor(banco_dados_jna_2025$menores_18_hogar,
                                                  levels = v65$levels,
                                                  labels = v65$labels)

# hogar gasto principal
banco_dados_jna_2025$hogar_gasto_principal= factor(banco_dados_jna_2025$hogar_gasto_principal,
                                             levels = v66$levels,
                                             labels = v66$labels)
# hogar renda mensal
banco_dados_jna_2025$hogar_renda_mensal= factor(banco_dados_jna_2025$hogar_renda_mensal,
                                                   levels = v67$levels,
                                                   labels = v67$labels)
# hogar dificuldade alimentacao
banco_dados_jna_2025$hogar_dificuldade_alimentacao= factor(banco_dados_jna_2025$hogar_dificuldade_alimentacao,
                                                levels = v68$levels,
                                                labels = v68$labels)

                                                  
# hogar razao dificuldade alimentacao
banco_dados_jna_2025$hogar_razao_dificuldade_alimentacao= factor(banco_dados_jna_2025$hogar_razao_dificuldade_alimentacao,
                                                levels = v69$levels,
                                                labels = v69$labels)

# hogar necessidade medica
banco_dados_jna_2025$hogar_necessidade_medica= factor(banco_dados_jna_2025$hogar_necessidade_medica,
                                                                 levels = v70$levels,
                                                                 labels = v70$labels)
# hogar necessidade medica ginecologica
banco_dados_jna_2025$hogar_necessidade_medica_ginecologica= factor(banco_dados_jna_2025$hogar_necessidade_medica_ginecologica,
                                                      levels = v71$levels,
                                                      labels = v71$labels)
# hogar necessidade medica pediatrica
banco_dados_jna_2025$hogar_necessidade_medica_pediatrica= factor(banco_dados_jna_2025$hogar_necessidade_medica_pediatrica,
                                                                   levels = v72$levels,
                                                                   labels = v72$labels)
# hogar necessidade medica neurologica
banco_dados_jna_2025$hogar_necessidade_medica_neurologica= factor(banco_dados_jna_2025$hogar_necessidade_medica_neurologica,
                                                                 levels = v73$levels,
                                                                 labels = v73$labels)
# hogar necessidade medica cardiologica
banco_dados_jna_2025$hogar_necessidade_medica_cardiologica= factor(banco_dados_jna_2025$hogar_necessidade_medica_cardiologica,
                                                                  levels = v74$levels,
                                                                  labels = v74$labels)
# hogar necessidade medica saude mental
banco_dados_jna_2025$hogar_necessidade_medica_saude_mental= factor(banco_dados_jna_2025$hogar_necessidade_medica_saude_mental,
                                                                   levels = v75$levels,
                                                                   labels = v75$labels)
# hogar necessidade medica nutricional
banco_dados_jna_2025$hogar_necessidade_medica_nutricional= factor(banco_dados_jna_2025$hogar_necessidade_medica_nutricional,
                                                                   levels = v76$levels,
                                                                   labels = v76$labels)
# hogar necessidade medica clinico geral
banco_dados_jna_2025$hogar_necessidade_medica_clinico_geral= factor(banco_dados_jna_2025$hogar_necessidade_medica_clinico_geral,
                                                                  levels = v77$levels,
                                                                  labels = v77$labels)
# hogar necessidade medica outro
banco_dados_jna_2025$hogar_necessidade_medica_outro= factor(banco_dados_jna_2025$hogar_necessidade_medica_outro,
                                                                    levels = v78$levels,
                                                                    labels = v78$labels)
# hogar necessidade medica ns
banco_dados_jna_2025$hogar_necessidade_medica_ns= factor(banco_dados_jna_2025$hogar_necessidade_medica_ns,
                                                            levels = v79$levels,
                                                            labels = v79$labels)
# hogar necessidade medica nr
banco_dados_jna_2025$hogar_necessidade_medica_nr= factor(banco_dados_jna_2025$hogar_necessidade_medica_nr,
                                                         levels = v80$levels,
                                                         labels = v80$labels)
# hogar dificuldade acesso saúde
banco_dados_jna_2025$hogar_dificuldade_acesso_saude= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude,
                                                         levels = v81$levels,
                                                         labels = v81$labels)
# hogar dificuldade acesso saúde, atendimento lento
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_lento= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_lento,
                                                           levels = v82$levels,
                                                           labels = v82$labels)
# hogar dificuldade acesso saúde, atendimento custo
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_custo= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_custo,
                                                                 levels = v83$levels,
                                                                 labels = v83$labels)
# hogar dificuldade acesso saúde, falta de documentacao
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_documentacao= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_documentacao,
                                                                 levels = v84$levels,
                                                                 labels = v84$labels)
# hogar dificuldade acesso saúde, unidades cheias
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_cheio= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_cheio,
                                                                        levels = v85$levels,
                                                                        labels = v85$labels)
# hogar dificuldade acesso saúde, falta especialista
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_especialista= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_especialista,
                                                                 levels = v86$levels,
                                                                 labels = v86$labels)
# hogar dificuldade acesso saúde, falta de informacao
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_informacao= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_informacao,
                                                                        levels = v87$levels,
                                                                        labels = v87$labels)
# hogar dificuldade acesso saúde, dificuldades com idioma
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_idioma= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_idioma,
                                                                      levels = v88$levels,
                                                                      labels = v88$labels)
# hogar dificuldade acesso saúde, discriminacao
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_discriminacao= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_discriminacao,
                                                                  levels = v89$levels,
                                                                  labels = v89$labels)

# hogar dificuldade acesso saúde, outro
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_outro= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_outro,
                                                                         levels = v90$levels,
                                                                         labels = v90$labels)
# hogar dificuldade acesso saúde, ns
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_ns= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_ns,
                                                                 levels = v91$levels,
                                                                 labels = v91$labels)
# hogar dificuldade acesso saúde, nr
banco_dados_jna_2025$hogar_dificuldade_acesso_saude_nr= factor(banco_dados_jna_2025$hogar_dificuldade_acesso_saude_nr,
                                                              levels = v92$levels,
                                                              labels = v92$labels)

# hogar diarreia, disenteria
banco_dados_jna_2025$hogar_diarreia_disenteria= factor(banco_dados_jna_2025$hogar_diarreia_disenteria,
                                                              levels = v93$levels,
                                                              labels = v93$labels)
# hogar acesso agua encanada
banco_dados_jna_2025$hogar_acesso_agua= factor(banco_dados_jna_2025$hogar_acesso_agua,
                                                       levels = v94$levels,
                                                       labels = v94$labels)

# hogar acesso saneamento basico
banco_dados_jna_2025$hogar_saneamento= factor(banco_dados_jna_2025$hogar_saneamento,
                                               levels = v95$levels,
                                               labels = v95$labels)
# hogar acesso coleta lixo
banco_dados_jna_2025$hogar_coleta_lixo_semanal= factor(banco_dados_jna_2025$hogar_coleta_lixo_semanal,
                                              levels = v96$levels,
                                              labels = v96$labels)

# hogar shelter utensilios de cozinha
banco_dados_jna_2025$hogar_abrigo_utensilios_cozinha= factor(banco_dados_jna_2025$hogar_abrigo_utensilios_cozinha,
                                                       levels = v97$levels,
                                                       labels = v97$labels)

# hogar abrigo cama
banco_dados_jna_2025$hogar_abrigo_cama= factor(banco_dados_jna_2025$hogar_abrigo_cama,
                                                             levels = v98$levels,
                                                             labels = v98$labels)
# hogar abrigo colchao
banco_dados_jna_2025$hogar_abrigo_colchao= factor(banco_dados_jna_2025$hogar_abrigo_colchao,
                                               levels = v99$levels,
                                               labels = v99$labels)
# hogar abrigo moveis
banco_dados_jna_2025$hogar_abrigo_moveis= factor(banco_dados_jna_2025$hogar_abrigo_moveis,
                                                  levels = v100$levels,
                                                  labels = v100$labels)
# hogar abrigo electrodomesticos
banco_dados_jna_2025$hogar_abrigo_electrodomesticos= factor(banco_dados_jna_2025$hogar_abrigo_electrodomesticos,
                                                 levels = v101$levels,
                                                 labels = v101$labels)
# hogar abrigo utensilios nr
banco_dados_jna_2025$hogar_abrigo_utensilios_nr= factor(banco_dados_jna_2025$hogar_abrigo_utensilios_nr,
                                                            levels = v102$levels,
                                                            labels = v102$labels)
# hogar moradia proximo mes
banco_dados_jna_2025$hogar_moradia_proximo_mes= factor(banco_dados_jna_2025$hogar_moradia_proximo_mes,
                                                        levels = v103$levels,
                                                        labels = v103$labels)
# hogar ja abrigado antes
banco_dados_jna_2025$hogar_ja_abrigo_antes= factor(banco_dados_jna_2025$hogar_ja_abrigo_antes,
                                                       levels = v104$levels,
                                                       labels = v104$labels)
# hogar discriminacao
banco_dados_jna_2025$hogar_discriminacao= factor(banco_dados_jna_2025$hogar_discriminacao,
                                                   levels = v105$levels,
                                                   labels = v105$labels)
# hogar uasc
banco_dados_jna_2025$hogar_uasc= factor(banco_dados_jna_2025$hogar_uasc,
                                                 levels = v106$levels,
                                                 labels = v106$labels)
# hogar trabalhando enganacao
banco_dados_jna_2025$hogar_trabalho_enganacao= factor(banco_dados_jna_2025$hogar_trabalho_enganacao,
                                        levels = v107$levels,
                                        labels = v107$labels)
# hogar trabalho sem pagamento
banco_dados_jna_2025$hogar_trabalho_sem_pagamento= factor(banco_dados_jna_2025$hogar_trabalho_sem_pagamento,
                                                      levels = v108$levels,
                                                      labels = v108$labels)
# seguranca gbv
banco_dados_jna_2025$seguranca_vbg= factor(banco_dados_jna_2025$seguranca_vbg,
                                                          levels = v109$levels,
                                                          labels = v109$labels)
# # participacao entrevista nome extra
# banco_dados_jna_2025$participacao_entrevista_extra= factor(banco_dados_jna_2025$participacao_entrevista_extra,
#                                            levels = v110$levels,
#                                            labels = v110$labels)

#rename  hogar id
banco_dados_jna_2025<-banco_dados_jna_2025 %>%
  rename(id_hogar=id)

# entrevistado dataset
entrevistado_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa", "entrevistado_venezuelano_a",
                    "maior_idade","chefe_hogar", "saber_responder", "entrevistado_genero", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "entrevistado_parentesco_chefe_hogar", "entrevistado_idade", "entrevistado_gravida_lactante", 
                    "entrevistado_gravida_servico_nutricao", "entrevistado_etnia", "entrevistado_deficiencia",
                    "entrevistado_deficiencia_fisica", "entrevistado_deficiencia_mental", "entrevistado_deficiencia_auditiva",
                    "entrevistado_deficiencia_visual", "entrevistado_deficiencia_autismo", "entrevistado_deficiencia_outro",
                    "entrevistado_documento_refugio_residencia", "entrevistado_cpf", "entrevistado_cnh", "entrevistado_sus",
                    "entrevistado_escola", "entrevistado_barreira_escola", "entrevistado_grau_escolar", "entrevistado_grau_escolar_completo",
                    "entrevistado_trabalhando", "entrevistado_contrato_trabalho", "entrevistado_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
                    ))

#rename entrevistado dataset


entrevistado_dataset<-entrevistado_dataset %>%
  add_column(entrevistado=1)

entrevistado_dataset<-entrevistado_dataset %>%
  rename(nacionalidade=entrevistado_venezuelano_a,
         genero=entrevistado_genero,
         idade_anos=entrevistado_idade,
         parentesco_chefe_hogar=entrevistado_parentesco_chefe_hogar,
         gravida_lactante=entrevistado_gravida_lactante,
         gravida_servico_nutricao=entrevistado_gravida_servico_nutricao,
         etnia=entrevistado_etnia,
         deficiencia=entrevistado_deficiencia,
         deficiencia_fisica=entrevistado_deficiencia_fisica,
         deficiencia_mental=entrevistado_deficiencia_mental,
         deficiencia_visual= entrevistado_deficiencia_visual,
         deficiencia_auditiva=entrevistado_deficiencia_auditiva,
         deficiencia_autismo = entrevistado_deficiencia_autismo,
         deficiencia_outro = entrevistado_deficiencia_outro,
         documento_refugio_residencia = entrevistado_documento_refugio_residencia,
         cpf=entrevistado_cpf,
         cnh=entrevistado_cnh,
         sus=entrevistado_sus,
         escola=entrevistado_escola,
         barreira_escola=entrevistado_barreira_escola,
         grau_escolar=entrevistado_grau_escolar,
         grau_escolar_completo=entrevistado_grau_escolar_completo,
         trabalhando=entrevistado_trabalhando,
         contrato_trabalho=entrevistado_contrato_trabalho,
         microempresa=entrevistado_microempresa,
         id_pessoa=entrevistado)



# pessoa1 dataset




pessoa1_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa1_genero", "pessoa1_parentesco", "pessoa1_idade_anos", "pessoa1_idade_meses",
                    "pessoa1_gravida_lactante", "pessoa1_gravida_servico_nutricao", 
                    "pessoa1_servico_nutricao_criancas", "pessoa1_etnia", "pessoa1_deficiencia",
                    "pessoa1_deficiencia_fisica", "pessoa1_deficiencia_mental",
                    "pessoa1_deficiencia_auditiva", "pessoa1_deficiencia_visual",
                    "pessoa1_deficiencia_autismo", "pessoa1_deficiencia_outro", "pessoa1_nacionalidade",
                    "pessoa1_documento_refugio_residencia", "pessoa1_cpf", "pessoa1_cnh",
                    "pessoa1_certidao_nascimento", "pessoa1_sus", "pessoa1_escola", "pessoa1_barreira_escola", 
                    "pessoa1_grau_escolar", "pessoa1_grau_escolar_completo", "pessoa1_trabalhando", 
                    "pessoa1_contrato_trabalho", "pessoa1_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa1 dataset

pessoa1_dataset<-pessoa1_dataset %>%
  add_column(pessoa1=2)

pessoa1_dataset<-pessoa1_dataset %>%
  rename(nacionalidade=pessoa1_nacionalidade,
         genero=pessoa1_genero,
         idade_anos=pessoa1_idade_anos,
         idade_meses=pessoa1_idade_meses,
         parentesco_chefe_hogar=pessoa1_parentesco,
         gravida_lactante=pessoa1_gravida_lactante,
         gravida_servico_nutricao=pessoa1_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa1_servico_nutricao_criancas,
         etnia=pessoa1_etnia,
         deficiencia=pessoa1_deficiencia,
         deficiencia_fisica=pessoa1_deficiencia_fisica,
         deficiencia_mental=pessoa1_deficiencia_mental,
         deficiencia_visual= pessoa1_deficiencia_visual,
         deficiencia_auditiva=pessoa1_deficiencia_auditiva,
         deficiencia_autismo = pessoa1_deficiencia_autismo,
         deficiencia_outro = pessoa1_deficiencia_outro,
         documento_refugio_residencia = pessoa1_documento_refugio_residencia,
         cpf=pessoa1_cpf,
         cnh=pessoa1_cnh,
         sus=pessoa1_sus,
         certidao_nascimento=pessoa1_certidao_nascimento,
         escola=pessoa1_escola,
         barreira_escola=pessoa1_barreira_escola,
         grau_escolar=pessoa1_grau_escolar,
         grau_escolar_completo=pessoa1_grau_escolar_completo,
         trabalhando=pessoa1_trabalhando,
         contrato_trabalho=pessoa1_contrato_trabalho,
         microempresa=pessoa1_microempresa,
        id_pessoa=pessoa1)
  
  
  
#pessoa 2 dataset



pessoa2_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa2_genero", "pessoa2_parentesco", "pessoa2_idade_anos", "pessoa2_idade_meses",
                    "pessoa2_gravida_lactante", "pessoa2_gravida_servico_nutricao", 
                    "pessoa2_servico_nutricao_criancas", "pessoa2_etnia", "pessoa2_deficiencia",
                    "pessoa2_deficiencia_fisica", "pessoa2_deficiencia_mental",
                    "pessoa2_deficiencia_auditiva", "pessoa2_deficiencia_visual",
                    "pessoa2_deficiencia_autismo", "pessoa2_deficiencia_outro", "pessoa2_nacionalidade",
                    "pessoa2_documento_refugio_residencia", "pessoa2_cpf", "pessoa2_cnh",
                    "pessoa2_certidao_nascimento", "pessoa2_sus", "pessoa2_escola", "pessoa2_barreira_escola", 
                    "pessoa2_grau_escolar", "pessoa2_grau_escolar_completo", "pessoa2_trabalhando", 
                    "pessoa2_contrato_trabalho", "pessoa2_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa2 dataset


pessoa2_dataset<-pessoa2_dataset %>%
  add_column(pessoa2=3)

pessoa2_dataset<-pessoa2_dataset %>%
  rename(nacionalidade=pessoa2_nacionalidade,
         genero=pessoa2_genero,
         idade_anos=pessoa2_idade_anos,
         idade_meses=pessoa2_idade_meses,
         parentesco_chefe_hogar=pessoa2_parentesco,
         gravida_lactante=pessoa2_gravida_lactante,
         gravida_servico_nutricao=pessoa2_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa2_servico_nutricao_criancas,
         etnia=pessoa2_etnia,
         deficiencia=pessoa2_deficiencia,
         deficiencia_fisica=pessoa2_deficiencia_fisica,
         deficiencia_mental=pessoa2_deficiencia_mental,
         deficiencia_visual= pessoa2_deficiencia_visual,
         deficiencia_auditiva=pessoa2_deficiencia_auditiva,
         deficiencia_autismo = pessoa2_deficiencia_autismo,
         deficiencia_outro = pessoa2_deficiencia_outro,
         documento_refugio_residencia = pessoa2_documento_refugio_residencia,
         cpf=pessoa2_cpf,
         cnh=pessoa2_cnh,
         sus=pessoa2_sus,
         certidao_nascimento=pessoa2_certidao_nascimento,
         escola=pessoa2_escola,
         barreira_escola=pessoa2_barreira_escola,
         grau_escolar=pessoa2_grau_escolar,
         grau_escolar_completo=pessoa2_grau_escolar_completo,
         trabalhando=pessoa2_trabalhando,
         contrato_trabalho=pessoa2_contrato_trabalho,
         microempresa=pessoa2_microempresa,
         id_pessoa=pessoa2)


# pessoa 3 dataset


pessoa3_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa3_genero", "pessoa3_parentesco", "pessoa3_idade_anos", "pessoa3_idade_meses",
                    "pessoa3_gravida_lactante", "pessoa3_gravida_servico_nutricao", 
                    "pessoa3_servico_nutricao_criancas", "pessoa3_etnia", "pessoa3_deficiencia",
                    "pessoa3_deficiencia_fisica", "pessoa3_deficiencia_mental",
                    "pessoa3_deficiencia_auditiva", "pessoa3_deficiencia_visual",
                    "pessoa3_deficiencia_autismo", "pessoa3_deficiencia_outro", "pessoa3_nacionalidade",
                    "pessoa3_documento_refugio_residencia", "pessoa3_cpf", "pessoa3_cnh",
                    "pessoa3_certidao_nascimento", "pessoa3_sus", "pessoa3_escola", "pessoa3_barreira_escola", 
                    "pessoa3_grau_escolar", "pessoa3_grau_escolar_completo", "pessoa3_trabalhando", 
                    "pessoa3_contrato_trabalho", "pessoa3_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa3 dataset

pessoa3_dataset<-pessoa3_dataset %>%
  add_column(pessoa3=4)


pessoa3_dataset<-pessoa3_dataset %>%
  rename(nacionalidade=pessoa3_nacionalidade,
         genero=pessoa3_genero,
         idade_anos=pessoa3_idade_anos,
         idade_meses=pessoa3_idade_meses,
         parentesco_chefe_hogar=pessoa3_parentesco,
         gravida_lactante=pessoa3_gravida_lactante,
         gravida_servico_nutricao=pessoa3_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa3_servico_nutricao_criancas,
         etnia=pessoa3_etnia,
         deficiencia=pessoa3_deficiencia,
         deficiencia_fisica=pessoa3_deficiencia_fisica,
         deficiencia_mental=pessoa3_deficiencia_mental,
         deficiencia_visual= pessoa3_deficiencia_visual,
         deficiencia_auditiva=pessoa3_deficiencia_auditiva,
         deficiencia_autismo = pessoa3_deficiencia_autismo,
         deficiencia_outro = pessoa3_deficiencia_outro,
         documento_refugio_residencia = pessoa3_documento_refugio_residencia,
         cpf=pessoa3_cpf,
         cnh=pessoa3_cnh,
         sus=pessoa3_sus,
         certidao_nascimento=pessoa3_certidao_nascimento,
         escola=pessoa3_escola,
         barreira_escola=pessoa3_barreira_escola,
         grau_escolar=pessoa3_grau_escolar,
         grau_escolar_completo=pessoa3_grau_escolar_completo,
         trabalhando=pessoa3_trabalhando,
         contrato_trabalho=pessoa3_contrato_trabalho,
         microempresa=pessoa3_microempresa,
         id_pessoa=pessoa3)

#pessoa 4 dataset



pessoa4_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa4_genero", "pessoa4_parentesco", "pessoa4_idade_anos", "pessoa4_idade_meses",
                    "pessoa4_gravida_lactante", "pessoa4_gravida_servico_nutricao", 
                    "pessoa4_servico_nutricao_criancas", "pessoa4_etnia", "pessoa4_deficiencia",
                    "pessoa4_deficiencia_fisica", "pessoa4_deficiencia_mental",
                    "pessoa4_deficiencia_auditiva", "pessoa4_deficiencia_visual",
                    "pessoa4_deficiencia_autismo", "pessoa4_deficiencia_outro", "pessoa4_nacionalidade",
                    "pessoa4_documento_refugio_residencia", "pessoa4_cpf", "pessoa4_cnh",
                    "pessoa4_certidao_nascimento", "pessoa4_sus", "pessoa4_escola", "pessoa4_barreira_escola", 
                    "pessoa4_grau_escolar", "pessoa4_grau_escolar_completo", "pessoa4_trabalhando", 
                    "pessoa4_contrato_trabalho", "pessoa4_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa4 dataset


pessoa4_dataset<-pessoa4_dataset %>%
  add_column(pessoa4=5)

pessoa4_dataset<-pessoa4_dataset %>%
  rename(nacionalidade=pessoa4_nacionalidade,
         genero=pessoa4_genero,
         idade_anos=pessoa4_idade_anos,
         idade_meses=pessoa4_idade_meses,
         parentesco_chefe_hogar=pessoa4_parentesco,
         gravida_lactante=pessoa4_gravida_lactante,
         gravida_servico_nutricao=pessoa4_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa4_servico_nutricao_criancas,
         etnia=pessoa4_etnia,
         deficiencia=pessoa4_deficiencia,
         deficiencia_fisica=pessoa4_deficiencia_fisica,
         deficiencia_mental=pessoa4_deficiencia_mental,
         deficiencia_visual= pessoa4_deficiencia_visual,
         deficiencia_auditiva=pessoa4_deficiencia_auditiva,
         deficiencia_autismo = pessoa4_deficiencia_autismo,
         deficiencia_outro = pessoa4_deficiencia_outro,
         documento_refugio_residencia = pessoa4_documento_refugio_residencia,
         cpf=pessoa4_cpf,
         cnh=pessoa4_cnh,
         sus=pessoa4_sus,
         certidao_nascimento=pessoa4_certidao_nascimento,
         escola=pessoa4_escola,
         barreira_escola=pessoa4_barreira_escola,
         grau_escolar=pessoa4_grau_escolar,
         grau_escolar_completo=pessoa4_grau_escolar_completo,
         trabalhando=pessoa4_trabalhando,
         contrato_trabalho=pessoa4_contrato_trabalho,
         microempresa=pessoa4_microempresa,
         id_pessoa=pessoa4)



#pessoa 5 dataset

pessoa5_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa5_genero", "pessoa5_parentesco", "pessoa5_idade_anos", "pessoa5_idade_meses",
                    "pessoa5_gravida_lactante", "pessoa5_gravida_servico_nutricao", 
                    "pessoa5_servico_nutricao_criancas", "pessoa5_etnia", "pessoa5_deficiencia",
                    "pessoa5_deficiencia_fisica", "pessoa5_deficiencia_mental",
                    "pessoa5_deficiencia_auditiva", "pessoa5_deficiencia_visual",
                    "pessoa5_deficiencia_autismo", "pessoa5_deficiencia_outro", "pessoa5_nacionalidade",
                    "pessoa5_documento_refugio_residencia", "pessoa5_cpf", "pessoa5_cnh",
                    "pessoa5_certidao_nascimento", "pessoa5_sus", "pessoa5_escola", "pessoa5_barreira_escola", 
                    "pessoa5_grau_escolar", "pessoa5_grau_escolar_completo", "pessoa5_trabalhando", 
                    "pessoa5_contrato_trabalho", "pessoa5_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa5 dataset

pessoa5_dataset<-pessoa5_dataset %>%
  add_column(pessoa5=6)



pessoa5_dataset<-pessoa5_dataset %>%
  rename(nacionalidade=pessoa5_nacionalidade,
         genero=pessoa5_genero,
         idade_anos=pessoa5_idade_anos,
         idade_meses=pessoa5_idade_meses,
         parentesco_chefe_hogar=pessoa5_parentesco,
         gravida_lactante=pessoa5_gravida_lactante,
         gravida_servico_nutricao=pessoa5_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa5_servico_nutricao_criancas,
         etnia=pessoa5_etnia,
         deficiencia=pessoa5_deficiencia,
         deficiencia_fisica=pessoa5_deficiencia_fisica,
         deficiencia_mental=pessoa5_deficiencia_mental,
         deficiencia_visual= pessoa5_deficiencia_visual,
         deficiencia_auditiva=pessoa5_deficiencia_auditiva,
         deficiencia_autismo = pessoa5_deficiencia_autismo,
         deficiencia_outro = pessoa5_deficiencia_outro,
         documento_refugio_residencia = pessoa5_documento_refugio_residencia,
         cpf=pessoa5_cpf,
         cnh=pessoa5_cnh,
         sus=pessoa5_sus,
         certidao_nascimento=pessoa5_certidao_nascimento,
         escola=pessoa5_escola,
         barreira_escola=pessoa5_barreira_escola,
         grau_escolar=pessoa5_grau_escolar,
         grau_escolar_completo=pessoa5_grau_escolar_completo,
         trabalhando=pessoa5_trabalhando,
         contrato_trabalho=pessoa5_contrato_trabalho,
         microempresa=pessoa5_microempresa,
         id_pessoa=pessoa5)


#pessoa6 dataset


pessoa6_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa6_genero", "pessoa6_parentesco", "pessoa6_idade_anos", "pessoa6_idade_meses",
                    "pessoa6_gravida_lactante", "pessoa6_gravida_servico_nutricao", 
                    "pessoa6_servico_nutricao_criancas", "pessoa6_etnia", "pessoa6_deficiencia",
                    "pessoa6_deficiencia_fisica", "pessoa6_deficiencia_mental",
                    "pessoa6_deficiencia_auditiva", "pessoa6_deficiencia_visual",
                    "pessoa6_deficiencia_autismo", "pessoa6_deficiencia_outro", "pessoa6_nacionalidade",
                    "pessoa6_documento_refugio_residencia", "pessoa6_cpf", "pessoa6_cnh",
                    "pessoa6_certidao_nascimento", "pessoa6_sus", "pessoa6_escola", "pessoa6_barreira_escola", 
                    "pessoa6_grau_escolar", "pessoa6_grau_escolar_completo", "pessoa6_trabalhando", 
                    "pessoa6_contrato_trabalho", "pessoa6_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa6 dataset

pessoa6_dataset<-pessoa6_dataset %>%
  add_column(pessoa6=7)

pessoa6_dataset<-pessoa6_dataset %>%
  rename(nacionalidade=pessoa6_nacionalidade,
         genero=pessoa6_genero,
         idade_anos=pessoa6_idade_anos,
         idade_meses=pessoa6_idade_meses,
         parentesco_chefe_hogar=pessoa6_parentesco,
         gravida_lactante=pessoa6_gravida_lactante,
         gravida_servico_nutricao=pessoa6_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa6_servico_nutricao_criancas,
         etnia=pessoa6_etnia,
         deficiencia=pessoa6_deficiencia,
         deficiencia_fisica=pessoa6_deficiencia_fisica,
         deficiencia_mental=pessoa6_deficiencia_mental,
         deficiencia_visual= pessoa6_deficiencia_visual,
         deficiencia_auditiva=pessoa6_deficiencia_auditiva,
         deficiencia_autismo = pessoa6_deficiencia_autismo,
         deficiencia_outro = pessoa6_deficiencia_outro,
         documento_refugio_residencia = pessoa6_documento_refugio_residencia,
         cpf=pessoa6_cpf,
         cnh=pessoa6_cnh,
         sus=pessoa6_sus,
         certidao_nascimento=pessoa6_certidao_nascimento,
         escola=pessoa6_escola,
         barreira_escola=pessoa6_barreira_escola,
         grau_escolar=pessoa6_grau_escolar,
         grau_escolar_completo=pessoa6_grau_escolar_completo,
         trabalhando=pessoa6_trabalhando,
         contrato_trabalho=pessoa6_contrato_trabalho,
         microempresa=pessoa6_microempresa,
         id_pessoa=pessoa6)

#pessoa7 dataset


pessoa7_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa7_genero", "pessoa7_parentesco", "pessoa7_idade_anos", "pessoa7_idade_meses",
                    "pessoa7_gravida_lactante", "pessoa7_gravida_servico_nutricao", 
                    "pessoa7_servico_nutricao_criancas", "pessoa7_etnia", "pessoa7_deficiencia",
                    "pessoa7_deficiencia_fisica", "pessoa7_deficiencia_mental",
                    "pessoa7_deficiencia_auditiva", "pessoa7_deficiencia_visual",
                    "pessoa7_deficiencia_autismo", "pessoa7_deficiencia_outro", "pessoa7_nacionalidade",
                    "pessoa7_documento_refugio_residencia", "pessoa7_cpf", "pessoa7_cnh",
                    "pessoa7_certidao_nascimento", "pessoa7_sus", "pessoa7_escola", "pessoa7_barreira_escola", 
                    "pessoa7_grau_escolar", "pessoa7_grau_escolar_completo", "pessoa7_trabalhando", 
                    "pessoa7_contrato_trabalho", "pessoa7_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa7 dataset

pessoa7_dataset<-pessoa7_dataset %>%
  add_column(pessoa7=8)

pessoa7_dataset<-pessoa7_dataset %>%
  rename(nacionalidade=pessoa7_nacionalidade,
         genero=pessoa7_genero,
         idade_anos=pessoa7_idade_anos,
         idade_meses=pessoa7_idade_meses,
         parentesco_chefe_hogar=pessoa7_parentesco,
         gravida_lactante=pessoa7_gravida_lactante,
         gravida_servico_nutricao=pessoa7_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa7_servico_nutricao_criancas,
         etnia=pessoa7_etnia,
         deficiencia=pessoa7_deficiencia,
         deficiencia_fisica=pessoa7_deficiencia_fisica,
         deficiencia_mental=pessoa7_deficiencia_mental,
         deficiencia_visual= pessoa7_deficiencia_visual,
         deficiencia_auditiva=pessoa7_deficiencia_auditiva,
         deficiencia_autismo = pessoa7_deficiencia_autismo,
         deficiencia_outro = pessoa7_deficiencia_outro,
         documento_refugio_residencia = pessoa7_documento_refugio_residencia,
         cpf=pessoa7_cpf,
         cnh=pessoa7_cnh,
         sus=pessoa7_sus,
         certidao_nascimento=pessoa7_certidao_nascimento,
         escola=pessoa7_escola,
         barreira_escola=pessoa7_barreira_escola,
         grau_escolar=pessoa7_grau_escolar,
         grau_escolar_completo=pessoa7_grau_escolar_completo,
         trabalhando=pessoa7_trabalhando,
         contrato_trabalho=pessoa7_contrato_trabalho,
         microempresa=pessoa7_microempresa,
         id_pessoa=pessoa7)

#pessoa8 dataset


pessoa8_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa8_genero", "pessoa8_parentesco", "pessoa8_idade_anos", "pessoa8_idade_meses",
                    "pessoa8_gravida_lactante", "pessoa8_gravida_servico_nutricao", 
                    "pessoa8_servico_nutricao_criancas", "pessoa8_etnia", "pessoa8_deficiencia",
                    "pessoa8_deficiencia_fisica", "pessoa8_deficiencia_mental",
                    "pessoa8_deficiencia_auditiva", "pessoa8_deficiencia_visual",
                    "pessoa8_deficiencia_autismo", "pessoa8_deficiencia_outro", "pessoa8_nacionalidade",
                    "pessoa8_documento_refugio_residencia", "pessoa8_cpf", "pessoa8_cnh",
                    "pessoa8_certidao_nascimento", "pessoa8_sus", "pessoa8_escola", "pessoa8_barreira_escola", 
                    "pessoa8_grau_escolar", "pessoa8_grau_escolar_completo", "pessoa8_trabalhando", 
                    "pessoa8_contrato_trabalho", "pessoa8_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa8 dataset


pessoa8_dataset<-pessoa8_dataset %>%
  add_column(pessoa8=9)

pessoa8_dataset<-pessoa8_dataset %>%
  rename(nacionalidade=pessoa8_nacionalidade,
         genero=pessoa8_genero,
         idade_anos=pessoa8_idade_anos,
         idade_meses=pessoa8_idade_meses,
         parentesco_chefe_hogar=pessoa8_parentesco,
         gravida_lactante=pessoa8_gravida_lactante,
         gravida_servico_nutricao=pessoa8_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa8_servico_nutricao_criancas,
         etnia=pessoa8_etnia,
         deficiencia=pessoa8_deficiencia,
         deficiencia_fisica=pessoa8_deficiencia_fisica,
         deficiencia_mental=pessoa8_deficiencia_mental,
         deficiencia_visual= pessoa8_deficiencia_visual,
         deficiencia_auditiva=pessoa8_deficiencia_auditiva,
         deficiencia_autismo = pessoa8_deficiencia_autismo,
         deficiencia_outro = pessoa8_deficiencia_outro,
         documento_refugio_residencia = pessoa8_documento_refugio_residencia,
         cpf=pessoa8_cpf,
         cnh=pessoa8_cnh,
         sus=pessoa8_sus,
         certidao_nascimento=pessoa8_certidao_nascimento,
         escola=pessoa8_escola,
         barreira_escola=pessoa8_barreira_escola,
         grau_escolar=pessoa8_grau_escolar,
         grau_escolar_completo=pessoa8_grau_escolar_completo,
         trabalhando=pessoa8_trabalhando,
         contrato_trabalho=pessoa8_contrato_trabalho,
         microempresa=pessoa8_microempresa,
         id_pessoa=pessoa8)


#pessoa9 dataset



pessoa9_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa9_genero", "pessoa9_parentesco", "pessoa9_idade_anos", "pessoa9_idade_meses",
                    "pessoa9_gravida_lactante", "pessoa9_gravida_servico_nutricao", 
                    "pessoa9_servico_nutricao_criancas", "pessoa9_etnia", "pessoa9_deficiencia",
                    "pessoa9_deficiencia_fisica", "pessoa9_deficiencia_mental",
                    "pessoa9_deficiencia_auditiva", "pessoa9_deficiencia_visual",
                    "pessoa9_deficiencia_autismo", "pessoa9_deficiencia_outro", "pessoa9_nacionalidade",
                    "pessoa9_documento_refugio_residencia", "pessoa9_cpf", "pessoa9_cnh",
                    "pessoa9_certidao_nascimento", "pessoa9_sus", "pessoa9_escola", "pessoa9_barreira_escola", 
                    "pessoa9_grau_escolar", "pessoa9_grau_escolar_completo", "pessoa9_trabalhando", 
                    "pessoa9_contrato_trabalho", "pessoa9_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa9 dataset
pessoa9_dataset<-pessoa9_dataset %>%
  add_column(pessoa9=10)


pessoa9_dataset<-pessoa9_dataset %>%
  rename(nacionalidade=pessoa9_nacionalidade,
         genero=pessoa9_genero,
         idade_anos=pessoa9_idade_anos,
         idade_meses=pessoa9_idade_meses,
         parentesco_chefe_hogar=pessoa9_parentesco,
         gravida_lactante=pessoa9_gravida_lactante,
         gravida_servico_nutricao=pessoa9_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa9_servico_nutricao_criancas,
         etnia=pessoa9_etnia,
         deficiencia=pessoa9_deficiencia,
         deficiencia_fisica=pessoa9_deficiencia_fisica,
         deficiencia_mental=pessoa9_deficiencia_mental,
         deficiencia_visual= pessoa9_deficiencia_visual,
         deficiencia_auditiva=pessoa9_deficiencia_auditiva,
         deficiencia_autismo = pessoa9_deficiencia_autismo,
         deficiencia_outro = pessoa9_deficiencia_outro,
         documento_refugio_residencia = pessoa9_documento_refugio_residencia,
         cpf=pessoa9_cpf,
         cnh=pessoa9_cnh,
         sus=pessoa9_sus,
         certidao_nascimento=pessoa9_certidao_nascimento,
         escola=pessoa9_escola,
         barreira_escola=pessoa9_barreira_escola,
         grau_escolar=pessoa9_grau_escolar,
         grau_escolar_completo=pessoa9_grau_escolar_completo,
         trabalhando=pessoa9_trabalhando,
         contrato_trabalho=pessoa9_contrato_trabalho,
         microempresa=pessoa9_microempresa,
         id_pessoa=pessoa9)


#pessoa10 dataset


pessoa10_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa10_genero", "pessoa10_parentesco", "pessoa10_idade_anos", "pessoa10_idade_meses",
                    "pessoa10_gravida_lactante", "pessoa10_gravida_servico_nutricao", 
                    "pessoa10_servico_nutricao_criancas", "pessoa10_etnia", "pessoa10_deficiencia",
                    "pessoa10_deficiencia_fisica", "pessoa10_deficiencia_mental",
                    "pessoa10_deficiencia_auditiva", "pessoa10_deficiencia_visual",
                    "pessoa10_deficiencia_autismo", "pessoa10_deficiencia_outro", "pessoa10_nacionalidade",
                    "pessoa10_documento_refugio_residencia", "pessoa10_cpf", "pessoa10_cnh",
                    "pessoa10_certidao_nascimento", "pessoa10_sus", "pessoa10_escola", "pessoa10_barreira_escola", 
                    "pessoa10_grau_escolar", "pessoa10_grau_escolar_completo", "pessoa10_trabalhando", 
                    "pessoa10_contrato_trabalho", "pessoa10_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa10 dataset

pessoa10_dataset<-pessoa10_dataset %>%
  add_column(pessoa10=11)

pessoa10_dataset<-pessoa10_dataset %>%
  rename(nacionalidade=pessoa10_nacionalidade,
         genero=pessoa10_genero,
         idade_anos=pessoa10_idade_anos,
         idade_meses=pessoa10_idade_meses,
         parentesco_chefe_hogar=pessoa10_parentesco,
         gravida_lactante=pessoa10_gravida_lactante,
         gravida_servico_nutricao=pessoa10_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa10_servico_nutricao_criancas,
         etnia=pessoa10_etnia,
         deficiencia=pessoa10_deficiencia,
         deficiencia_fisica=pessoa10_deficiencia_fisica,
         deficiencia_mental=pessoa10_deficiencia_mental,
         deficiencia_visual= pessoa10_deficiencia_visual,
         deficiencia_auditiva=pessoa10_deficiencia_auditiva,
         deficiencia_autismo = pessoa10_deficiencia_autismo,
         deficiencia_outro = pessoa10_deficiencia_outro,
         documento_refugio_residencia = pessoa10_documento_refugio_residencia,
         cpf=pessoa10_cpf,
         cnh=pessoa10_cnh,
         sus=pessoa10_sus,
         certidao_nascimento=pessoa10_certidao_nascimento,
         escola=pessoa10_escola,
         barreira_escola=pessoa10_barreira_escola,
         grau_escolar=pessoa10_grau_escolar,
         grau_escolar_completo=pessoa10_grau_escolar_completo,
         trabalhando=pessoa10_trabalhando,
         contrato_trabalho=pessoa10_contrato_trabalho,
         microempresa=pessoa10_microempresa,
         id_pessoa=pessoa10)


#pessoa11 dataset


pessoa11_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa11_genero", "pessoa11_parentesco", "pessoa11_idade_anos", "pessoa11_idade_meses",
                    "pessoa11_gravida_lactante", "pessoa11_gravida_servico_nutricao", 
                    "pessoa11_servico_nutricao_criancas", "pessoa11_etnia", "pessoa11_deficiencia",
                    "pessoa11_deficiencia_fisica", "pessoa11_deficiencia_mental",
                    "pessoa11_deficiencia_auditiva", "pessoa11_deficiencia_visual",
                    "pessoa11_deficiencia_autismo", "pessoa11_deficiencia_outro", "pessoa11_nacionalidade",
                    "pessoa11_documento_refugio_residencia", "pessoa11_cpf", "pessoa11_cnh",
                    "pessoa11_certidao_nascimento", "pessoa11_sus", "pessoa11_escola", "pessoa11_barreira_escola", 
                    "pessoa11_grau_escolar", "pessoa11_grau_escolar_completo", "pessoa11_trabalhando", 
                    "pessoa11_contrato_trabalho", "pessoa11_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa11 dataset

pessoa11_dataset<-pessoa11_dataset %>%
  add_column(pessoa11=12)


pessoa11_dataset<-pessoa11_dataset %>%
  rename(nacionalidade=pessoa11_nacionalidade,
         genero=pessoa11_genero,
         idade_anos=pessoa11_idade_anos,
         idade_meses=pessoa11_idade_meses,
         parentesco_chefe_hogar=pessoa11_parentesco,
         gravida_lactante=pessoa11_gravida_lactante,
         gravida_servico_nutricao=pessoa11_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa11_servico_nutricao_criancas,
         etnia=pessoa11_etnia,
         deficiencia=pessoa11_deficiencia,
         deficiencia_fisica=pessoa11_deficiencia_fisica,
         deficiencia_mental=pessoa11_deficiencia_mental,
         deficiencia_visual= pessoa11_deficiencia_visual,
         deficiencia_auditiva=pessoa11_deficiencia_auditiva,
         deficiencia_autismo = pessoa11_deficiencia_autismo,
         deficiencia_outro = pessoa11_deficiencia_outro,
         documento_refugio_residencia = pessoa11_documento_refugio_residencia,
         cpf=pessoa11_cpf,
         cnh=pessoa11_cnh,
         sus=pessoa11_sus,
         certidao_nascimento=pessoa11_certidao_nascimento,
         escola=pessoa11_escola,
         barreira_escola=pessoa11_barreira_escola,
         grau_escolar=pessoa11_grau_escolar,
         grau_escolar_completo=pessoa11_grau_escolar_completo,
         trabalhando=pessoa11_trabalhando,
         contrato_trabalho=pessoa11_contrato_trabalho,
         microempresa=pessoa11_microempresa,
         id_pessoa=pessoa11)


#pessoa12 dataset


pessoa12_dataset <- 
  subset(banco_dados_jna_2025,
         select = c("id_hogar","data_inicio", "data_fim","base_contatos","ja_participou_pesquisa",
                    "maior_idade","chefe_hogar", "saber_responder", "municipio",
                    "estado", "ano_chegada_brasil", "interiorizado_a", "intencao_movimento", "tamanho_hogar",
                    "pessoa12_genero", "pessoa12_parentesco", "pessoa12_idade_anos", "pessoa12_idade_meses",
                    "pessoa12_gravida_lactante", "pessoa12_gravida_servico_nutricao", 
                    "pessoa12_servico_nutricao_criancas", "pessoa12_etnia", "pessoa12_deficiencia",
                    "pessoa12_deficiencia_fisica", "pessoa12_deficiencia_mental",
                    "pessoa12_deficiencia_auditiva", "pessoa12_deficiencia_visual",
                    "pessoa12_deficiencia_autismo", "pessoa12_deficiencia_outro", "pessoa12_nacionalidade",
                    "pessoa12_documento_refugio_residencia", "pessoa12_cpf", "pessoa12_cnh",
                    "pessoa12_certidao_nascimento", "pessoa12_sus", "pessoa12_escola", "pessoa12_barreira_escola", 
                    "pessoa12_grau_escolar", "pessoa12_grau_escolar_completo", "pessoa12_trabalhando", 
                    "pessoa12_contrato_trabalho", "pessoa12_microempresa",
                    "menores_18_hogar", "hogar_gasto_principal","hogar_renda_mensal","hogar_dificuldade_alimentacao",
                    "hogar_razao_dificuldade_alimentacao","hogar_necessidade_medica","hogar_necessidade_medica_ginecologica",
                    "hogar_necessidade_medica_pediatrica", "hogar_necessidade_medica_neurologica", "hogar_necessidade_medica_cardiologica",
                    "hogar_necessidade_medica_saude_mental", "hogar_necessidade_medica_nutricional", "hogar_necessidade_medica_clinico_geral",
                    "hogar_necessidade_medica_outro", "hogar_necessidade_medica_ns", "hogar_necessidade_medica_nr",
                    "hogar_dificuldade_acesso_saude", "hogar_dificuldade_acesso_saude_lento", "hogar_dificuldade_acesso_saude_custo",
                    "hogar_dificuldade_acesso_saude_documentacao", "hogar_dificuldade_acesso_saude_cheio", "hogar_dificuldade_acesso_saude_especialista",
                    "hogar_dificuldade_acesso_saude_informacao", "hogar_dificuldade_acesso_saude_idioma",
                    "hogar_dificuldade_acesso_saude_discriminacao", "hogar_dificuldade_acesso_saude_outro",
                    "hogar_dificuldade_acesso_saude_ns", "hogar_dificuldade_acesso_saude_nr", 
                    "hogar_diarreia_disenteria", "hogar_acesso_agua", "hogar_saneamento", "hogar_coleta_lixo_semanal",
                    "hogar_abrigo_utensilios_cozinha", "hogar_abrigo_cama","hogar_abrigo_colchao", "hogar_abrigo_moveis",
                    "hogar_abrigo_electrodomesticos", "hogar_abrigo_utensilios_nr", "hogar_moradia_proximo_mes", 
                    "hogar_ja_abrigo_antes", "hogar_discriminacao", "hogar_uasc", "hogar_trabalho_enganacao", "hogar_trabalho_sem_pagamento",
                    "seguranca_vbg"
         ))

#rename pessoa12 dataset

pessoa12_dataset<-pessoa12_dataset %>%
  add_column(pessoa12=13)

pessoa12_dataset<-pessoa12_dataset %>%
  rename(nacionalidade=pessoa12_nacionalidade,
         genero=pessoa12_genero,
         idade_anos=pessoa12_idade_anos,
         idade_meses=pessoa12_idade_meses,
         parentesco_chefe_hogar=pessoa12_parentesco,
         gravida_lactante=pessoa12_gravida_lactante,
         gravida_servico_nutricao=pessoa12_gravida_servico_nutricao,
         servico_nutricao_criancas=pessoa12_servico_nutricao_criancas,
         etnia=pessoa12_etnia,
         deficiencia=pessoa12_deficiencia,
         deficiencia_fisica=pessoa12_deficiencia_fisica,
         deficiencia_mental=pessoa12_deficiencia_mental,
         deficiencia_visual= pessoa12_deficiencia_visual,
         deficiencia_auditiva=pessoa12_deficiencia_auditiva,
         deficiencia_autismo = pessoa12_deficiencia_autismo,
         deficiencia_outro = pessoa12_deficiencia_outro,
         documento_refugio_residencia = pessoa12_documento_refugio_residencia,
         cpf=pessoa12_cpf,
         cnh=pessoa12_cnh,
         sus=pessoa12_sus,
         certidao_nascimento=pessoa12_certidao_nascimento,
         escola=pessoa12_escola,
         barreira_escola=pessoa12_barreira_escola,
         grau_escolar=pessoa12_grau_escolar,
         grau_escolar_completo=pessoa12_grau_escolar_completo,
         trabalhando=pessoa12_trabalhando,
         contrato_trabalho=pessoa12_contrato_trabalho,
         microempresa=pessoa12_microempresa,
         id_pessoa=pessoa12)



# merged <- bind_rows(entrevistado_dataset, pessoa1_dataset, pessoa2_dataset,
#                     pessoa3_dataset, pessoa4_dataset, pessoa5_dataset,
#                     pessoa6_dataset, pessoa7_dataset, pessoa8_dataset,
#                     pessoa9_dataset, pessoa10_dataset, pessoa11_dataset,
#                     pessoa12_dataset)

# === EMPILHAMENTO CORRIGIDO (ÚNICA ALTERAÇÃO) ================================
# Mantém todo o resto do seu arquivo intacto.
# Corrige o erro: "Can't combine ..1$idade_anos <double> and ..2$idade_anos <factor<...>>"

library(dplyr)
library(purrr)
library(readr)

# Lista os dataframes das pessoas

datasets_nomes <- c(
  "entrevistado_dataset",
  "pessoa1_dataset", "pessoa2_dataset", "pessoa3_dataset",
  "pessoa4_dataset", "pessoa5_dataset", "pessoa6_dataset",
  "pessoa7_dataset", "pessoa8_dataset", "pessoa9_dataset",
  "pessoa10_dataset", "pessoa11_dataset", "pessoa12_dataset"
)

# Recupera os objetos por nome (que já foram criados)
datasets_list <- mget(datasets_nomes, inherits = TRUE)

# Função de saneamento: se existir 'idade_anos', converte p/ numérico de forma segura
coerce_idade_anos <- function(df) {
  if ("idade_anos" %in% names(df)) {
    df <- df %>%
      mutate(
        # converte factor/character -> numeric, deixando NA quando não for número
        idade_anos = parse_double(as.character(idade_anos), na = c("", "NA", "NaN"))
      )
  }
  df
}

# Aplica saneamento em todos os dataframes
datasets_list_fixed <- map(datasets_list, coerce_idade_anos)

# Opcional: alinhar colunas (caso algum df tenha coluna a menos/mais)
todas_cols <- Reduce(union, lapply(datasets_list_fixed, names))
datasets_list_fixed <- map(
  datasets_list_fixed,
  ~ .x %>%
    add_column(!!!setNames(
      rep(list(NA), length(setdiff(todas_cols, names(.x)))),
      setdiff(todas_cols, names(.x))
    )) %>%
    select(all_of(todas_cols))
)

# Empilha sem conflito de tipos
merged <- bind_rows(datasets_list_fixed)

# Sanidade: confirme que idade_anos ficou numérica (se existir)
if ("idade_anos" %in% names(merged)) {
  stopifnot(is.numeric(merged$idade_anos))
}
# === FIM DO EMPILHAMENTO CORRIGIDO ===========================================



#create index column for person

merged <- merged %>%                                        # Create ID by group
  group_by(id_hogar) %>%
  dplyr::mutate(id_pessoa_hogar = cur_group_id())


#remove blanks from genero
merged <- merged[!(is.na(merged$genero)), ]

#transform idade meses em anos

library(dplyr)
merged<- merged %>%
  mutate(
    idade_meses = case_when(
      idade_meses==0|idade_meses==1|idade_meses==2|idade_meses==3|idade_meses==4
      |idade_meses==5|idade_meses==6|idade_meses==7|idade_meses==8|idade_meses==9
      |idade_meses==10|idade_meses==11  ~0,
      idade_meses==12|idade_meses==13|idade_meses==14
      |idade_meses==15|idade_meses==16|idade_meses==17|idade_meses==18|idade_meses==19
      |idade_meses==22|idade_meses==23~1,
      idade_meses==24|idade_meses==25|idade_meses==26
      |idade_meses==27|idade_meses==28|idade_meses==29|idade_meses==30|idade_meses==31
      |idade_meses==32|idade_meses==33|idade_meses==34|idade_meses==35~2,
      idade_meses==36 |idade_meses==37|idade_meses==38|idade_meses==39|idade_meses==40|
      idade_meses==41|idade_meses==42|idade_meses==43|idade_meses==44|idade_meses==45|idade_meses==46
      |idade_meses==47~3,
      idade_meses==48|idade_meses==49|idade_meses==50|idade_meses==51
      |idade_meses==52|idade_meses==53|idade_meses==54 ~ 4
 ))

#put together idade_anos and idade_meses

merged <- transform(merged, idade_anos_new = ifelse(!is.na(idade_anos), idade_anos, idade_meses))

# add região

library(dplyr)
merged<- merged%>%
  mutate(regiao = case_when(
      estado=="AMAZONAS"| estado=="RORAIMA"| estado=="AMAPA"
      | estado=="PARA"| estado=="TOCANTINS"| estado=="RONDONIA"
      | estado=="ACRE"~"Norte",
      estado=="MARANHAO"| estado=="PIAUI"| estado=="CEARA"
      | estado=="RIO GRANDE DO NORTE"| estado=="PERNAMBUCO"| estado=="PARAIBA"
      | estado=="SERGIPE"| estado=="ALAGOAS"| estado=="BAHIA"~"Nordeste",
      estado=="MATO GROSSO"| estado=="MATO GROSSO DO SUL"| estado=="DISTRITO FEDERAL"
      | estado=="GOIAS"~"Centro-Oeste",
      estado=="SAO PAULO"| estado=="RIO DE JANEIRO"| estado=="ESPIRITO SANTO"
      | estado=="MINAS GERAIS"~"Sudeste",
       estado=="PARANA"| estado=="RIO GRANDE DO SUL"| estado=="SANTA CATARINA"
      ~"Sul",
      estado==-99~"NR",
      estado==-88~"NS"
  ))

#transform etnia

library(dplyr)
merged<- merged %>%
  mutate(etnia_grupos = case_when(
    etnia=="Afrodescendente"~"Afrodescendente",
    etnia=="Branco(a)"~"Branco(a)",
    etnia=="Indígena"~"Indígena",
    etnia=="Amarelo(a)"|etnia=="Amarelo(a)"~"Outra",
    etnia=="Amarelo (a)"|etnia=="Amarelo(a)"~"Outra",
    etnia=="Moreno(a)"|etnia=="Criollo (a)"|etnia=="Mestiço(a)"
    |etnia=="Pardo(a)"|etnia=="Trigueño(a)"|etnia=="Trigueño (a)"~"Pardo(a)",
    etnia=="NS"~"NS",
    etnia=="NR"~"NR"
))

#transform capital

library(dplyr)
merged<- merged %>%
  mutate(capital = case_when(
    municipio=="MANAUS (AMAZONAS)"| municipio=="BOA VISTA (RORAIMA)"| municipio=="MACAPA (AMAPA)"
    | municipio=="BELEM (PARA)"| municipio=="PALMAS (TOCANTINS)"| municipio=="PORTO VELHO(RONDONIA)"
    | municipio=="RIO BRANCO (ACRE)"| municipio=="SAO LUIS (MARANHAO)"| municipio=="TERESINA (PIAUI)"| municipio=="FORTALEZA (CEARA)"
    | municipio=="NATAL (RIO GRANDE DO NORTE)"| municipio=="RECIFE (PERNAMBUCO)"| municipio=="JOAO PESSOA (PARAIBA)"
    | municipio=="ARACAJU (SERGIPE)"| municipio=="MACEIO (ALAGOAS)"| municipio=="SALVADOR (BAHIA)"
    |municipio=="CUIABA (MATO GROSSO)"| municipio=="CAMPO GRANDE (MATO GROSSO DO SUL)"| municipio=="BRASILIA (DISTRITO FEDERAL)"
    | municipio=="GOIANIA (GOIAS)"
    | municipio=="SAO PAULO (SAO PAULO)"| municipio=="RIO DE JANEIRO (RIO DE JANEIRO)"| municipio=="VITORIA (ESPIRITO SANTO)"
    | municipio=="BELO HORIZONTE (MINAS GERAIS)"
    |municipio=="CURITIBA (PARANA)"| municipio=="PORTO ALEGRE (RIO GRANDE DO SUL)"| municipio=="FLORIANOPOLIS (SANTA CATARINA)"
    ~"Sim",
    municipio==-99~"NR",
    municipio==-88~"NS",
    TRUE~"Não"
))


# add RR and AM

library(dplyr)
merged<- merged %>%
  mutate(RR_AM = case_when(
    estado=="RORAIMA"| estado=="AMAZONAS"~"Sim",
    TRUE~"Não"
))


# add RR

library(dplyr)
merged<- merged %>%
  mutate(RR= case_when(
    estado=="RORAIMA"~"Sim",
    TRUE~"Não"
 ))

# add AM

library(dplyr)
merged<- merged %>%
  mutate(AM= case_when(
    estado=="AMAZONAS"~"Sim",
    TRUE~"Não"
))


# add idade grupos

library(dplyr)
merged<- merged %>%
  mutate(idade_grupos = case_when(
    idade_anos<=5~"0-5",
    idade_anos>=6& idade_anos<=11~"6-11",
    idade_anos>=12& idade_anos<=17~"12-17",
    idade_anos>=18& idade_anos<=30~"18-30",
    idade_anos>=31& idade_anos<=45~"31-45",
    idade_anos>=46& idade_anos<=65~"46-65",
    idade_anos>=66~"Mais que 65",
    TRUE~""
  ))


# add idade grupos numeros

library(dplyr)
merged<- merged %>%
  mutate(idade_grupos_numero = case_when(
    idade_grupos=="0-5"~1,
    idade_grupos=="6-11"~2,
    idade_grupos=="12-17"~3,
    idade_grupos=="18-30"~4,
    idade_grupos=="31-45"~5,
    idade_grupos=="46-65"~6,
    idade_grupos=="Mais que 65"~7,
    TRUE~0
  ))

# chefe hogar genero

library(dplyr)
merged<- merged %>%
  mutate(chefe_hogar_genero = case_when(
    parentesco_chefe_hogar=="Chefe de hogar"& genero=="Feminino"~"Chefe de hogar feminino",
    parentesco_chefe_hogar=="Chefe de hogar"& genero=="Masculino"~"Chefe de hogar masculino",
    parentesco_chefe_hogar=="Chefe de hogar"& genero=="Outro"~"Chefe de hogar outro",
    parentesco_chefe_hogar=="Chefe de hogar"& genero=="NR"~"Chefe de hogar NR",
    TRUE~"Não é chefe de hogar"
  ))

# composicao familiar

library(dplyr)
merged<- merged %>%
  mutate(composicao_familiar = case_when(
    parentesco_chefe_hogar=="Avô/avó"|parentesco_chefe_hogar=="Genro/nora"|
      parentesco_chefe_hogar=="Irmão(a)"|parentesco_chefe_hogar=="Neto(a)/bisneto(a)"|
      parentesco_chefe_hogar=="Outro parente"|parentesco_chefe_hogar=="Sogro(a)"~"Fora do núcleo familiar",
      parentesco_chefe_hogar=="Chefe de hogar"~"Chefe de hogar",
      parentesco_chefe_hogar=="Conjuge"~"Conjuge",
      parentesco_chefe_hogar=="Pai, mãe, padrasto, madrasta"~"Pai, mãe, padrasto, madrasta",
    parentesco_chefe_hogar== "Um filho(a)/enteado(a)"~"Um filho(a)/enteado(a)",
    parentesco_chefe_hogar== "Outra pessoa (não familiar)"~"Outra pessoa (não familiar)",
    parentesco_chefe_hogar== "NR"~"NR"
  ))

#in_formalidade das pessoas entrevistadas

merged<- merged %>%
  mutate(
    pessoa_in_formalidade = case_when(
      trabalhando=="Sí, es empleado/trabajador" & contrato_trabalho=="No"|trabalhando=="Sí, es emprendedor" & microempresa=="No"~'Informal',
      trabalhando=="Sí, es empleado/trabajador" & contrato_trabalho=="Sí"|trabalhando=="Sí, es emprendedor" & microempresa=="Sí"~'Formal',
      TRUE ~ 'Não se aplica'
    ))

#in_formalidade emprego das pessoas entrevistadas

merged<- merged %>%
  mutate(
    pessoa_in_formalidade_emprego = case_when(
      trabalhando=="Sí, es empleado/trabajador" & contrato_trabalho=="No"~'Informal',
      trabalhando=="Sí, es empleado/trabajador" & contrato_trabalho=="Sí"~'Formal',
      TRUE ~ 'Não se aplica'
    ))

#in_formalidade emprendedores das pessoas entrevistadas

merged<- merged %>%
  mutate(
    pessoa_in_formalidade_emprendedor = case_when(
      trabalhando=="Sí, es emprendedor" & microempresa=="No"~'Informal',
      trabalhando=="Sí, es emprendedor" & microempresa=="Sí"~'Formal',
      TRUE ~ 'Não se aplica'
    ))

# id_hogares sobra

library(dplyr)
merged<- merged %>%
  mutate(id_hogar_sobras = case_when(
    id_hogar==271326| id_hogar==284912| id_hogar==275200
    | id_hogar==277996| id_hogar==105927| id_hogar==225766
    | id_hogar==288767| id_hogar==205056| id_hogar==171269| id_hogar==303730
    | id_hogar==53249| id_hogar==285484| id_hogar==131689
    | id_hogar==260621| id_hogar==281913| id_hogar==28374
    |id_hogar==162545| id_hogar==260918| id_hogar==103728
    | id_hogar==74313| id_hogar==279805| id_hogar==298568| id_hogar==281471
    | id_hogar==66287
    |id_hogar==142685| id_hogar==279595| id_hogar==170450
    |id_hogar==32695| id_hogar==255030| id_hogar==247728
    |id_hogar==291706| id_hogar==250584~"Sim",
    TRUE~"Não"
  ))

# criancas de 0-2

library(dplyr)
merged<- merged %>%
  mutate(crianca_0_2 = case_when(
    idade_anos_new==0|idade_anos_new==1
    |idade_anos_new==2~"Sim",
    TRUE~"Não"
  ))

# criancas de 0-5

library(dplyr)
merged<- merged %>%
  mutate(crianca_0_5 = case_when(
    idade_anos_new==0|idade_anos_new==1
    |idade_anos_new==2|idade_anos_new==3
    |idade_anos_new==4|idade_anos_new==5~"Sim",
    TRUE~"Não"
  ))

# criancas de 6-17


library(dplyr)
merged<- merged %>%
  mutate(crianca_6_17 = case_when(
    idade_anos_new==6|idade_anos_new==7
    |idade_anos_new==8|idade_anos_new==9
    |idade_anos_new==10|idade_anos_new==11
    |idade_anos_new==12
    |idade_anos_new==13|idade_anos_new==14
    |idade_anos_new==15|idade_anos_new==16
    |idade_anos_new==17~"Sim",
    TRUE~"Não"
  ))



# hogar com criancas e adol 6-17

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_crianca_6_17= any(idade_anos_new==6
                                 |idade_anos_new==7
                                 |idade_anos_new==8
                                 |idade_anos_new==9
                                 |idade_anos_new==10
                                 |idade_anos_new==11
                                 |idade_anos_new==12
                                 |idade_anos_new==13
                                 |idade_anos_new==14
                                 |idade_anos_new==15
                                 |idade_anos_new==16
                                 |idade_anos_new==17))


# https://stackoverflow.com/questions/31576879/finding-if-boolean-is-ever-true-by-groups-in-r

#hogar com crianca 0-2

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_crianca_0_2 = any(idade_anos_new <=2))

#hogar com crianca 0-5

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_crianca_0_5 = any(idade_anos_new <=5))

#hogar com crianca ate 

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_crianca_6_17= any(idade_anos_new==6
                                 |idade_anos_new==7
                                 |idade_anos_new==8
                                 |idade_anos_new==9
                                 |idade_anos_new==10
                                 |idade_anos_new==11
                                 |idade_anos_new==12
                                 |idade_anos_new==13
                                 |idade_anos_new==14
                                 |idade_anos_new==15
                                 |idade_anos_new==16
                                 |idade_anos_new==17))

#hogar com lactante ou gravida

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_lactante_gravida = any(gravida_lactante =="Sim, gravida"|
                                      gravida_lactante == "Sim, amamentando/pós-parto"|
                                      gravida_lactante == "Sim, gravida e amamentando (ou pós-parto)"))

#hogar com indigena

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_indigena = any(etnia_grupos =="Indígena"))

#hogar com afrodescendente

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_afrodescendente = any(etnia_grupos =="Afrodescendente"))

# hogar chefiado por mulher

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefiado_mulher = any(chefe_hogar_genero =="Chefe de hogar feminino"))

# hogar com criancas e adol

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_criancas_adol = any(idade_anos_new <=17))


# hogar com conjuge

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_com_conjuge = any(composicao_familiar == "Conjuge"))

# hogar sem conjuge, criancas e adol e chefe de hogar mulher

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefiado_mulher_monoparental = any(hogar_com_conjuge == "FALSE" &
                                                  hogar_criancas_adol == "TRUE"&
                                                  hogar_chefiado_mulher=="TRUE"))

# hogar com pessoa deficiente

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_deficiencia = any(deficiencia == "Sim"))


# hogar com chefe de hogar na informalidade (emprego informal ou emprendedorimso informal)

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefe_hogar_informal = any(composicao_familiar == "Chefe de hogar" &
                                          trabalhando == "Sí, es empleado/trabajador"&
                                          contrato_trabalho=="No"|
                                          composicao_familiar == "Chefe de hogar" &
                                          trabalhando == "Sí, es emprendedor"&
                                          microempresa=="No" 
                                          ))


# hogar com chefe de hogar na formalidade (emprego formal ou emprendedorimso formal)


merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefe_hogar_formal = any(composicao_familiar == "Chefe de hogar" &
                                            trabalhando == "Sí, es empleado/trabajador"&
                                            contrato_trabalho=="Sí"|
                                            composicao_familiar == "Chefe de hogar" &
                                            trabalhando == "Sí, es emprendedor"&
                                            microempresa=="Sí" 
  ))


# hogar com chefe de hogar com emprego formal

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefe_hogar_emprego_formal = any(composicao_familiar == "Chefe de hogar" &
                                                trabalhando == "Sí, es empleado/trabajador"&
                                                contrato_trabalho=="Sí"
  ))


# hogar com chefe de hogar com emprendedor formal

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefe_hogar_emprendedor_formal = any(composicao_familiar == "Chefe de hogar" &
                                                    trabalhando == "Sí, es emprendedor"&
                                                    microempresa=="Sí" 
  ))




# hogar com chefe de hogar com emprego informal

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefe_hogar_emprego_informal = any(composicao_familiar == "Chefe de hogar" &
                                                  trabalhando == "Sí, es empleado/trabajador"&
                                                  contrato_trabalho=="No"
  ))



# hogar com chefe de hogar com eempreendedor informal


merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_chefe_hogar_emprendedor_informal = any(composicao_familiar == "Chefe de hogar" &
                                                      trabalhando == "Sí, es emprendedor"&
                                                      microempresa=="No" 
  ))


#hogares onde tem pessoas que exercem atividade remunerada

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(hogar_pessoas_atividade_remunerada = any(trabalhando == "Independiente/ autónomo" |
                                                    trabalhando == "Profesional independiente/ autónomo"|
                                                    trabalhando== "Sí, como aprendiz" |
                                                    trabalhando =="Sí, es empleado/trabajador"|
                                                    trabalhando == "Sí, es emprendedor"|
                                                    trabalhando == "Sí, es profesional independiente/autónomo"|
                                                    trabalhando == "Sí, trabaja al diario"
  ))

#pessoas que exercem atividade remunerada

merged <- merged %>%
  group_by(id_hogar) %>%
  mutate(pessoas_atividade_remunerada = case_when(trabalhando == "Independiente/ autónomo" |
                                                    trabalhando == "Profesional independiente/ autónomo"|
                                                    trabalhando== "Sí, como aprendiz" |
                                                    trabalhando =="Sí, es empleado/trabajador"|
                                                    trabalhando == "Sí, es emprendedor"|
                                                    trabalhando == "Sí, es profesional independiente/autónomo"|
                                                    trabalhando == "Sí, trabaja al diario"
                                                  ~"Sim",
                                                  TRUE~"Não"
  )
  )

# ==== BLOCO FINAL (corrigido): pessoas em formato longo, idades e merges ====
library(dplyr)
library(tidyr)
library(readr)
library(stringr)

# 1) Chave do domicílio (usa a primeira coluna disponível; se não houver, cria)
key_candidates <- c("id_hogar","id_hogar_tirar2","uuid","KEY","_uuid","id")
available_keys <- intersect(key_candidates, names(banco_dados_jna_2025))
if (length(available_keys) == 0) {
  banco_dados_jna_2025 <- banco_dados_jna_2025 %>% mutate(hh_id = row_number())
} else {
  k <- available_keys[1]
  banco_dados_jna_2025 <- banco_dados_jna_2025 %>% mutate(hh_id = as.character(.data[[k]]))
}

# 2) Pessoas em formato longo (forçando tipos para character para evitar erro)
cols_pessoas <- names(banco_dados_jna_2025)[str_detect(names(banco_dados_jna_2025), "^pessoa\\d+_")]
pessoas_long <- banco_dados_jna_2025 %>%
  select(hh_id, any_of(cols_pessoas)) %>%
  pivot_longer(
    cols = -hh_id,
    names_to   = c("idx","campo"),
    names_pattern = "pessoa(\\d+)_(.*)",
    values_to  = "valor",
    values_ptypes     = list(valor = character()),
    values_transform  = list(valor = as.character)
  ) %>%
  mutate(idx = as.integer(idx)) %>%
  pivot_wider(names_from = "campo", values_from = "valor")

# 3) Idade calculada (anos) e faixas etárias
pessoas_long <- pessoas_long %>%
  mutate(
    idade_anos_n  = suppressWarnings(parse_number(as.character(idade_anos))),
    idade_meses_n = suppressWarnings(parse_number(as.character(idade_meses))),
    idade_anos_calc = dplyr::coalesce(idade_anos_n, idade_meses_n/12)
  ) %>%
  mutate(
    faixa_idade = case_when(
      !is.na(idade_anos_calc) & idade_anos_calc < 3   ~ "0-2",
      !is.na(idade_anos_calc) & idade_anos_calc < 6   ~ "3-5",
      !is.na(idade_anos_calc) & idade_anos_calc < 18  ~ "6-17",
      !is.na(idade_anos_calc) & idade_anos_calc < 60  ~ "18-59",
      !is.na(idade_anos_calc) & idade_anos_calc >= 60 ~ "60+",
      TRUE ~ NA_character_
    )
  )

# 4) Agregados do lar (flags)
hh_flags <- pessoas_long %>%
  group_by(hh_id) %>%
  summarise(
    crianca_6_17                = any(faixa_idade == "6-17", na.rm = TRUE),
    hogar_crianca_6_17          = crianca_6_17,
    hogar_crianca_0_2           = any(faixa_idade == "0-2", na.rm = TRUE),
    hogar_crianca_0_5           = any(faixa_idade %in% c("0-2","3-5"), na.rm = TRUE),
    hogar_criancas_adol         = any(!is.na(idade_anos_calc) & idade_anos_calc < 18, na.rm = TRUE),
    
    hogar_lactante_gravida      = any(str_to_lower(as.character(gravida_lactante)) %in% c("sim","yes","true","1"), na.rm = TRUE),
    hogar_deficiencia           = any(str_to_lower(as.character(deficiencia)) %in% c("sim","yes","true","1"), na.rm = TRUE),
    hogar_indigena              = any(str_detect(str_to_lower(as.character(etnia)), "indig"), na.rm = TRUE),
    hogar_afrodescendente       = any(str_detect(str_to_lower(as.character(etnia)), "afro"), na.rm = TRUE),
    
    # Chefe mulher / cônjuge
    chefe_mulher                = any(str_detect(str_to_lower(as.character(parentesco)), "chefe") &
                                        str_detect(str_to_lower(as.character(genero)), "fem"), na.rm = TRUE),
    hogar_chefiado_mulher       = chefe_mulher,
    hogar_com_conjuge           = any(str_detect(str_to_lower(as.character(parentesco)),
                                                 "c[oô]njuge|companheir"), na.rm = TRUE),
    hogar_chefiado_mulher_monoparental = chefe_mulher & !hogar_com_conjuge,
    
    # Emprego do chefe (heurística)
    chefe_trabalha              = any(str_detect(str_to_lower(as.character(parentesco)), "chefe") &
                                        str_to_lower(as.character(trabalhando)) %in% c("sim","yes","true","1"), na.rm = TRUE),
    chefe_contrato_txt          = paste(unique(na.omit(str_to_lower(as.character(contrato_trabalho)))), collapse=";"),
    chefe_microempresa          = any(str_detect(str_to_lower(as.character(parentesco)), "chefe") &
                                        str_to_lower(as.character(microempresa)) %in% c("sim","yes","true","1"), na.rm = TRUE),
    hogar_chefe_hogar_emprego_formal      = chefe_trabalha & str_detect(chefe_contrato_txt, "formal"),
    hogar_chefe_hogar_emprego_informal    = chefe_trabalha & str_detect(chefe_contrato_txt, "informal"),
    hogar_chefe_hogar_emprendedor_formal  = chefe_microempresa & str_detect(chefe_contrato_txt, "formal"),
    hogar_chefe_hogar_emprendedor_informal= chefe_microempresa & (str_detect(chefe_contrato_txt, "informal") | chefe_contrato_txt == ""),
    
    # Pessoas com atividade remunerada
    qtde_pessoas_atividade_remunerada     = sum(str_to_lower(as.character(trabalhando)) %in% c("sim","yes","true","1"), na.rm = TRUE),
    hogar_pessoas_atividade_remunerada    = qtde_pessoas_atividade_remunerada > 0,
    pessoas_atividade_remunerada          = qtde_pessoas_atividade_remunerada > 0
  ) %>%
  ungroup()

# 5) Converter flags para "Sim/Não"
cols_bool <- c(
  "crianca_6_17","hogar_crianca_6_17","hogar_crianca_0_2","hogar_crianca_0_5",
  "hogar_lactante_gravida","hogar_indigena","hogar_afrodescendente","hogar_chefiado_mulher",
  "hogar_criancas_adol","hogar_com_conjuge","hogar_chefiado_mulher_monoparental",
  "hogar_deficiencia","hogar_chefe_hogar_emprego_formal","hogar_chefe_hogar_emprendedor_formal",
  "hogar_chefe_hogar_emprego_informal","hogar_chefe_hogar_emprendedor_informal",
  "hogar_pessoas_atividade_remunerada","pessoas_atividade_remunerada"
)
hh_flags <- hh_flags %>%
  mutate(across(all_of(intersect(cols_bool, names(.))),
                ~factor(ifelse(is.na(.), NA, ifelse(., "Sim","Não")), levels=c("Não","Sim"))))

# 6) Idades em formato wide (uma coluna por pessoa)
idade_wide <- pessoas_long %>%
  transmute(hh_id,
            col = paste0("pessoa", idx, "_idade_anos_calc"),
            idade_anos_calc) %>%
  pivot_wider(names_from = col, values_from = idade_anos_calc)

# 7) MERGES finais no banco principal
resultado_final <- banco_dados_jna_2025 %>%
  left_join(hh_flags,  by = "hh_id") %>%
  left_join(idade_wide, by = "hh_id")




write_xlsx(merged,"C:/Users/SILVADOS/OneDrive - UNHCR/Desktop/Desktop/R4V/R4V 2025/JNA 2025/dash JNA 2025/Base completa/merged_2025.xlsx")



