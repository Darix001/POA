import streamlit as st
import zipfile
import ujson
import io

from itertools import starmap
from test import pyxl_process as modify, utf_op

config :dict[str, dict[str, bool|dict[str, list[str]]]]

with utf_op("config.json") as configfile:
	config = ujson.load(configfile)

st.set_page_config(
	page_title="Corrector POA 2025/2026",
	page_icon="📂",	
	layout="centered"
	)



# File uploader (permite arrastrar o seleccionar manualmente)
uploaded_files = st.file_uploader("Seleccionar POA",
	type=("xlsx",),
	accept_multiple_files=True)

text = ''
outs = []


def download_button(label:str, filename:str, mime:str):
	st.download_button(
		label=label,
		data=buffer.getvalue(),
		file_name=filename,
		mime=mime,
		icon=":material/download:"
		)

if st.button("**Iniciar Proceso**", icon='▶', type="primary") and uploaded_files:
	with io.BytesIO() as buffer, st.spinner("Procesando POAs..."):
		if len(uploaded_files) == 1:
			file, = uploaded_files
			wb = modify(file, buffer, config)
			download_button(
				"Download POA", file.name,
				("application/vnd.openxmlformats-officedocument"\
				".spreadsheetml.sheet"),
				)
		else:
			with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
				for file in uploaded_files:
					with io.BytesIO() as out:
						modify(file, out, config)
						zipf.writestr(file.name, out.getvalue())

			download_button("Download POAs", "POAs.zip",
				"application/zip")

		st.success("POAs procesados exitosamente.")


st.markdown(":red[Ediciones Soportadas:] *2025, 2026*")


with st.container():
	st.subheader("**Configuración**")
	corrections = config["correct"]
	corrections.update(
		zip(corrections, starmap(st.checkbox, corrections.items()))
		)


if st.button("Guardar Configuración", type="secondary"):
	with utf_op("config.json", "w") as configfile:
		ujson.dump(config, configfile, indent=4, ensure_ascii=False)
	st.success("Nueva Configuración guardada.")