const readXlsxFile = require("read-excel-file/node");

const loadXlxs = (src) => {
  return readXlsxFile(src).then((rows) => rows);
};

const main = async () => {
  const src = "./public/calidad_datos.xlsx";
  // Carga de archivos
  const data = await loadXlxs(src);

  //Creacion de un objeto memo en el que se ingresen los errores de la informacion
  const errors = {};

  for (let i = 0; i < data[0].length; i++) {
    const currentColumn = data[0][i];
    errors[currentColumn] = {};
  }


  const keys = Object.keys(errors);

  // Clasificar la informacion cargada en la variable data
  for (let i = 1; i < data.length; i++) {
    for (let j = 0; j < data[0].length; j++) {
      const currentData = data[i][j];

      // Verificar si el dato de una columna es nulo
      if (currentData === null) {
        save("null", i, currentData, errors[keys[j]]);
        continue;
      }

      // Verificar si el dato de una columna tiene error (#REF!)
      if (currentData === "#ERROR_#REF!") {
        save("#REF!", i, currentData, errors[keys[j]]);
        continue;
      }

      // Verificar que se cumplen las reglas de cada columna (cedula, email, ...)
      if (!constraints(keys[j], currentData + "")) {
        save("invalid", i, currentData, errors[keys[j]]);
      }
    }
  }

  // Copia del objeto errors al objeto resumne
  const resume = JSON.parse(JSON.stringify(errors));

  // Cuenta los errores de cada campo y almacena el total
  let totalErrors = 0;
  for (const key in resume) {
    let count = 0;
    for (const data in resume[key]) {
      const currentErrors = resume[key][data].length;
      resume[key][data] = currentErrors;
      count += currentErrors;
    }

    totalErrors += count;
    resume[key].total = count;
  }

  console.log("Errors\n", errors);

  console.log("Errors of Identification Cards\n", errors["cedula_est"]);
  console.log("\nResume Errors\n", resume);
  console.log("Total Errors:", totalErrors);
};

const constraints = (column, string) => {
  switch (column) {
    case "codigo_est":
      return isStudentCode(string);
    case "correo_est":
      return isEmail(string);
    case "cedula_est":
      return isIdentificationCard(string);
    case "telefono_est":
      return isTelephoneNumber(string);
    default:
      return true;
  }
};

const isStudentCode = (code) => {
    code = "" + code;
    return code.length === 11;
  };

const isTelephoneNumber = (number) => {
  number = number + "";
  const regex = /^[2][0-9]{6}$/;
  return regex.test(number);
};

const isIdentificationCard = (identificationCard) => {
  if (
    typeof identificationCard == "string" &&
    identificationCard.length == 10 &&
    /^\d+$/.test(identificationCard)
  ) {
    const digitos = identificationCard.split("").map(Number);
    const codigo_provincia = digitos[0] * 10 + digitos[1];

    //if (codigo_provincia >= 1 && (codigo_provincia <= 24 || codigo_provincia == 30) && digitos[2] < 6) {

    if (
      codigo_provincia >= 1 &&
      (codigo_provincia <= 24 || codigo_provincia == 30)
    ) {
      const digito_verificador = digitos.pop();

      const digito_calculado =
        digitos.reduce(function (valorPrevio, valorActual, indice) {
          return (
            valorPrevio -
            ((valorActual * (2 - (indice % 2))) % 9) -
            (valorActual == 9) * 9
          );
        }, 1000) % 10;
      return digito_calculado === digito_verificador;
    }
  }
  return false;
};

const isEmail = (email) => {
  const regex =
    /^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$/;

  return regex.test(email.toLowerCase());
};

// Guarda un valor en el objeto memo que es recibido por parámetro
const save = (type, row, data, memo) => {
  if (memo[type] === undefined) {
    memo[type] = [];
  }

  memo[type].push({ row, data });
};


main();
