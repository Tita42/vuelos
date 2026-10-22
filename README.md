Rellenar mejor después:)

Para correr lo de sacar los datos modelos hay que: 

En Chrome, abrir la pagina donde aparecen los vuelos del dia
Apretar F12 e ir a la pestaña Console.
Abrir descargar_vuelos_ber.js con cualquier editor de texto, copiar todo y pegarlo en la consola. Apretar Enter.
Chrome puede bloquear el pegado y pedirte que escribas allow pasting. Escribirlo, apretar Enter y volver a pegar. 
Va a aparecer el avance página por página. Tarda uno o dos minutos. Al terminar se descarga vuelos_ber_api.json.

Descargar el archivo donde esta la tabla con los vuelos y correr 

python completar_aviones_ber.py vuelos_BER_2026-10-22_y_29.xlsx vuelos_ber_api.json

