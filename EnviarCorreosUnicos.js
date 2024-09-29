function enviarCorreosUnicos() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Obtenemos los datos de las celdas y los separamos por comas
    var nombres = sheet.getRange("A1").getValue().split(","); 
    var saludos = sheet.getRange("A2").getValue().split(","); 
    var asuntos = sheet.getRange("A3").getValue().split(","); 
    var mensajes = sheet.getRange("A4").getValue().split(","); 
    var despedidas = sheet.getRange("A5").getValue().split(","); 
  
    // Inicializamos un Array vacio para almacenar las combinaciones unicas de correos
    var correosUnicos = [];
    
    // Creamos combinaciones unicas de correos utilizando bucles FOR anidados
    for (let i = 0; i < nombres.length; i++) { 
      for (let j = 0; j < saludos.length; j++) { 
        for (let k = 0; k < asuntos.length; k++) { 
          for (let l = 0; l < mensajes.length; l++) { 
            for (let m = 0; m < despedidas.length; m++) { 
              
              // Formamos un nuevo objeto de correo con la combinacion actual
              const CORREO = {
                // Obtenemos los datos y eliminamos espacios con el metodo trim( ).
                nombre: nombres[i].trim(), 
                saludo: saludos[j].trim(), 
                asunto: asuntos[k].trim(),
                mensaje: mensajes[l].trim(),
                despedida: despedidas[m].trim(),
              };
  
              // Agregamos la combinacion actual al Array de correos unicos
              correosUnicos.push(CORREO);
            }
          }
        }
      }
    }
  
    // Filtramos correos unicos en caso de que haya combinaciones repetidas
    /*
     *  1. Utilizamos Array.from para llamar a Set, que es una estructura de datos unicos.
     *     Con ello, nos aseguramos de que no haya duplicados. Con "map", convertimos cada objeto
     *     del Array correosUnicos en una cadena de texto utilizando "JSON.stringify". 
     *     Una vez se comprueban los duplicados, volvemos a convertir las cadenas de texto a objetos
     *     con "JSON.parse".
     */
    const CORREOS_UNICOS_FINALES = Array.from(new Set(correosUnicos.map(JSON.stringify))).map(JSON.parse);
    
    // Limitamos el numero de correos a 30 con slice( ).
    var correosParaEnviar = CORREOS_UNICOS_FINALES.slice(0, 30);
  
    // Enviamos los correos utilizando el Array CORREOS_UNICOS_FINALES
    // correo => --- Funcion anonima, pertenece a forEach.
    correosParaEnviar.forEach(correo => {
      const MAIL_DESTINO = "INTRODUCE_EL_CORREO_DESTINO_AQUI";
      // Creamos el cuerpo del correo utilizando plantilla de texto
      const MENSAJE = `${correo.saludo} ${correo.nombre},\n\n${correo.mensaje}\n\n${correo.despedida}`;
      
      // Enviamos el correo usando MailApp
      MailApp.sendEmail({
        to: MAIL_DESTINO, 
        subject: correo.asunto, 
        body: MENSAJE, 
      });
    });
    
    Logger.log("Correos enviados: " + correosParaEnviar.length);
  }