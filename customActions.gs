function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Crea el menú principal "Actions"
  ui.createMenu('Actions')
    // Añade un botón "Get Groups" que ejecutará la función getGroups
    .addItem('Get All Groups', 'getAllGroups')
    .addItem('Empty Groups', 'emptyGroup')
    .addItem('Add Member to Group', 'addToGroup')
    .addItem('Create Users', 'createUser')
    .addItem('Create Groups', 'createGroup')
    .addItem('Delete Groups', 'deleteGroup')
    .addItem('Copy Groups', 'copyGroups')
    .addItem('Suspend Users', 'suspendUsers')
    .addToUi();
}


function getAllGroups() {
  // Limpia cualquier dato previo en la columna A
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('A2:A').clear();
  
  // Obtén todos los grupos del dominio
  var pageToken, page;
  var i = 2; // La fila donde empezar a listar
  do {
    page = AdminDirectory.Groups.list({
      customer: 'my_customer',
      maxResults: 200,
      pageToken: pageToken
    });
    var groups = page.groups;
    if (groups) {
      for (var j = 0; j < groups.length; j++) {
        sheet.getRange(i, 1).setValue(groups[j].email);
        i++;
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
  SpreadsheetApp.getUi().alert('All groups have been listed.');
}


function emptyGroup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  

  // Itera sobre cada grupo listado en la columna A desde la celda A2 en adelante
  for (var i = 2; i <= lastRow; i++) {
    var groupEmail = sheet.getRange(i, 1).getValue();
    
    if (groupEmail) {
      try {
        // Obtiene todos los miembros del grupo
        var members = AdminDirectory.Members.list(groupEmail).members;
        
        if (members && members.length > 0) {
          // Elimina cada usuario del grupo
          for (var j = 0; j < members.length; j++) {
            AdminDirectory.Members.remove(groupEmail, members[j].email);
          }
        }
        
        // Opcional: Marca la fila como "procesada" (por ejemplo, con un color o texto)
        sheet.getRange(i, 2).setValue('Emptied');
        
      } catch (e) {
        
        SpreadsheetApp.getUi().alert('Error processing group ' + groupEmail + ': ' + e.message);
      }
    } else {
      Logger.log("No se encontró correo electrónico en la fila: " + i);
    }
  }
  
  SpreadsheetApp.getUi().alert('All specified groups have been emptied.');
  Logger.log("Proceso completado.");
}


function addToGroup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Itera sobre cada fila a partir de la segunda fila
  for (var i = 2; i <= lastRow; i++) {
    var email = sheet.getRange(i, 1).getValue();
    var groupEmail = sheet.getRange(i, 2).getValue();
    
    if (email && groupEmail) {
      try {
        // Añade el usuario al grupo
        AdminDirectory.Members.insert({
          email: email,
          role: "MEMBER" // "MEMBER" por defecto, puede ser "OWNER" o "MANAGER"
        }, groupEmail);
        
        // Marca como añadido en la columna C
        sheet.getRange(i, 3).setValue('Added');
        
      } catch (e) {
        // Si hay un error, escribe el mensaje de error en la columna C
        sheet.getRange(i, 3).setValue(e.message);
      }
    } else {
      // Si no hay datos en alguna celda de la fila, marca como "Datos faltantes"
      sheet.getRange(i, 3).setValue('Datos faltantes');
    }
  }
}



function createUser() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var totalUsers = lastRow - 1; // Número total de usuarios a procesar
  var userData = sheet.getRange(2, 1, totalUsers, 5).getValues(); // Obtiene los datos de la hoja (A2:E)

  for (var i = 0; i < userData.length; i++) {
    var firstName = userData[i][0]; // Columna A: Nombre Gsuite
    var lastName = userData[i][1];  // Columna B: Apellidos Gsuite
    var email = userData[i][2];     // Columna C: Email
    var groupEmail = userData[i][3];// Columna D: Email Grupo
    var password = userData[i][4];  // Columna E: Contraseña

    if (firstName && lastName && email && password) {
      try {
        // Crea un objeto user con los detalles del nuevo usuario
        var user = {
          primaryEmail: email,
          name: {
            givenName: firstName,
            familyName: lastName
          },
          password: password,
          changePasswordAtNextLogin: true // Obliga al usuario a cambiar la contraseña en el primer inicio de sesión
        };

        // Asigna a la unidad organizativa según el dominio del correo
        if (email.endsWith(".al@iespoligonosur.org")) {
          user.orgUnitPath = '/Alumnado';
        } else {
          user.orgUnitPath = '/Profesorado'; // O cualquier otra unidad organizativa predeterminada
        }

        // Intenta crear el usuario en Google Workspace
        try {
          AdminDirectory.Users.insert(user);
          sheet.getRange(i + 2, 6).setValue('Created'); // Columna F: Resultado
        } catch (userError) {
          if (userError.message.includes("Entity already exists")) {
            sheet.getRange(i + 2, 6).setValue('Already exists');
          } else {
            throw userError;
          }
        }

        // Si el usuario pertenece a "Alumnado", intenta añadirlo al grupo especificado
        if (user.orgUnitPath === '/Alumnado' && groupEmail) {
          try {
            AdminDirectory.Members.insert({
              email: email,
              role: "MEMBER"
            }, groupEmail);
            sheet.getRange(i + 2, 7).setValue('Added to Group'); // Columna G: Resultado de la adición al grupo
          } catch (groupError) {
            sheet.getRange(i + 2, 7).setValue(groupError.message);
          }
        } else {
          // Si no pertenece a "Alumnado", escribe la unidad organizativa en la columna G
          sheet.getRange(i + 2, 7).setValue(user.orgUnitPath);
        }

      } catch (e) {
        sheet.getRange(i + 2, 6).setValue(e.message); // Columna F: Error
      }
    } else {
      sheet.getRange(i + 2, 6).setValue('Datos faltantes'); // Columna F: Datos faltantes
    }

    // Actualiza el porcentaje de usuarios procesados en la celda J1
    var percentage = ((i + 1) / totalUsers) * 100;
    sheet.getRange('J1').setValue(percentage.toFixed(2) + '%');
  }
}




function createGroup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Itera sobre cada fila a partir de la segunda fila
  for (var i = 2; i <= lastRow; i++) {
    var groupName = sheet.getRange(i, 1).getValue();   // Columna A: Nombre del grupo
    var groupEmail = sheet.getRange(i, 2).getValue();  // Columna B: Email del grupo
    var description = sheet.getRange(i, 3).getValue(); // Columna C: Descripción del grupo
    
    if (groupName && groupEmail) {
      try {
        // Crea un objeto group con los detalles del nuevo grupo
        var group = {
          name: groupName,
          email: groupEmail,
          description: description || '' // Si no hay descripción, se deja vacío
        };
        
        // Crea el grupo en Google Workspace
        AdminDirectory.Groups.insert(group);
        
        // Marca como creado en la columna D
        sheet.getRange(i, 4).setValue('Created');
        
      } catch (e) {
        // Si hay un error, escribe el mensaje de error en la columna D
        sheet.getRange(i, 4).setValue(e.message);
      }
    } else {
      // Si faltan datos en alguna celda de la fila, marca como "Datos faltantes"
      sheet.getRange(i, 4).setValue('Datos faltantes');
    }
  }
}


function deleteGroup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Itera sobre cada fila a partir de la segunda fila
  for (var i = 2; i <= lastRow; i++) {
    var groupEmail = sheet.getRange(i, 1).getValue(); // Columna A: Email del grupo
    
    if (groupEmail) {
      try {
        // Elimina el grupo en Google Workspace usando el email del grupo
        AdminDirectory.Groups.remove(groupEmail);
        
        // Marca como eliminado en la columna B
        sheet.getRange(i, 2).setValue('Deleted');
        
      } catch (e) {
        // Si hay un error, escribe el mensaje de error en la columna B
        sheet.getRange(i, 2).setValue(e.message);
      }
    } else {
      // Si el email del grupo falta, marca como "Datos faltantes"
      sheet.getRange(i, 2).setValue('Datos faltantes');
    }
  }
}


function copyGroups() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Obtener los correos electrónicos de usuario1 y usuario2
  var user1Email = sheet.getRange("A2").getValue(); // Usuario 1
  var user2Email = sheet.getRange("B2").getValue(); // Usuario 2
  
  if (!user1Email || !user2Email) {
    SpreadsheetApp.getUi().alert('Ambos correos electrónicos deben estar especificados en las celdas A2 y B2.');
    return;
  }

  try {
    // Obtener la lista de grupos a los que pertenece usuario1
    var groups = AdminDirectory.Groups.list({
      userKey: user1Email
    }).groups;

    if (!groups || groups.length === 0) {
      sheet.getRange("C2").setValue("No se encontraron grupos para " + user1Email);
      return;
    }

    var outputGroups = [];
    var outputResults = [];

    // Iterar sobre cada grupo y agregar usuario2
    for (var i = 0; i < groups.length; i++) {
      var groupEmail = groups[i].email;
      outputGroups.push([groupEmail]); // Añadir el grupo a la columna C

      try {
        AdminDirectory.Members.insert({
          email: user2Email,
          role: "MEMBER"
        }, groupEmail);
        
        outputResults.push(['Added']); // Indica que se añadió correctamente en la columna D
      } catch (e) {
        outputResults.push([e.message]); // Guarda el error en la columna D si no se pudo añadir
      }
    }

    // Escribir los resultados en las columnas C y D
    sheet.getRange(2, 3, outputGroups.length, 1).setValues(outputGroups); // Columna C
    sheet.getRange(2, 4, outputResults.length, 1).setValues(outputResults); // Columna D

  } catch (e) {
    sheet.getRange("C2").setValue("Error: " + e.message);
  }
}


function suspendUsers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues(); // Obtener todos los correos electrónicos desde la columna A
  var output = []; // Array para almacenar los resultados
  var processedCount = 0; // Contador de usuarios procesados

  for (var i = 0; i < data.length; i++) {
    var userEmail = data[i][0];

    if (userEmail) {
      try {
        // Suspender al usuario
        AdminDirectory.Users.update({
          suspended: true
        }, userEmail);

        output.push(['Suspended']); // Indicar que el usuario ha sido suspendido

      } catch (e) {
        output.push([e.message]); // Si ocurre un error, guardar el mensaje de error
      }
    } else {
      output.push(['Correo electrónico faltante']); // Si el email no está presente, indicar error
    }

    processedCount++;

    // Actualizar la celda E1 cada 5 usuarios procesados
    if (processedCount % 5 === 0) {
      sheet.getRange("E1").setValue(processedCount);
      SpreadsheetApp.flush(); // Aplicar los cambios inmediatamente
    }
  }

  // Escribir los resultados en la columna B
  sheet.getRange(2, 2, output.length, 1).setValues(output);

  // Actualizar el número total de usuarios procesados al final
  sheet.getRange("E1").setValue(processedCount);
}


function getTeachers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var groupEmail = "groupemail@domain.org"; // Reemplaza con el correo del grupo 
  var row = 2; // Comienza a escribir los datos desde la fila 2

  // Registra el tiempo de inicio
  var startTime = Date.now();

  try {
    // Obtiene todos los miembros del grupo "Grupo"
    var members = AdminDirectory.Members.list(groupEmail).members;

    if (members && members.length > 0) {
      var data = []; // Array para almacenar los datos a escribir en la hoja

      // Itera sobre cada miembro del grupo
      members.forEach(function(member) {
        // Verifica si el miembro es un usuario (puede haber otros grupos anidados)
        if (member.type === 'USER') {
          try {
            // Obtiene la información del usuario
            var user = AdminDirectory.Users.get(member.email);
            
            // Extrae el nombre, apellidos y correo del usuario
            var firstName = user.name.givenName;
            var lastName = user.name.familyName;
            var email = user.primaryEmail;

            // Almacena los datos en el array
            data.push([firstName, lastName, email]);
          } catch (userError) {
            Logger.log("Error al obtener datos del usuario: " + userError.message);
          }
        }
      });

      // Escribe todos los datos almacenados en la hoja de cálculo de una sola vez
      if (data.length > 0) {
        sheet.getRange(row, 1, data.length, 3).setValues(data);
      }
      
      // Forza la actualización de los datos en la hoja de cálculo
      SpreadsheetApp.flush();

    } else {
      Logger.log("No se encontraron miembros en el grupo 'Claustro'.");
    }
  } catch (e) {
    Logger.log("Error al obtener los miembros del grupo 'Claustro': " + e.message);
  }

  // Registra el tiempo de finalización
  var endTime = Date.now();
  
  // Calcula el tiempo total de ejecución en segundos
  var totalTime = ((endTime - startTime) / 1000).toFixed(2);

  // Muestra el tiempo total en la celda F1
  sheet.getRange('F1').setValue(totalTime + ' segundos');
}
