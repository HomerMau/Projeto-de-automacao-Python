import groupdocs_conversion_cloud

# Suas credenciais do GroupDocs
app_sid = "d5957486-1b7a-4800-bb49-c2b5c68df91f"
app_key = "cf43e2c8b8026cad5eee45367dba8ebf"

# Crie uma instância da API
api = groupdocs_conversion_cloud.InfoApi.from_keys(app_sid, app_key)

settings = groupdocs_conversion_cloud.ConvertSettings()
settings.file_path = "C://Users//garot//Downloads//BD1635-6.dwg"  # caminho do arquivo
settings.format = "pdf"
settings.output_path = "ConvertedFiles"  # caminho de saída

request = groupdocs_conversion_cloud.ConvertDocumentRequest(settings)
response = api.convert_document(request)

print("Arquivo convertido com sucesso. Verifique a pasta 'ConvertedFiles'.")
