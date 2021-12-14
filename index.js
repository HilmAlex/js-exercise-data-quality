const readXlsxFile = require("read-excel-file/node");

const loadXlxs = (src) => {
  return readXlsxFile(src).then((rows) => rows);
};

const main = async () => {
  const src = "./public/calidad_datos.xlsx";
  // Carga de archivos
  const data = await loadXlxs(src);

  //Creacion de un objeto memo en el que se ingresen los errores de la informacion
  const memo = {
    codigo_est: {},
    cedula_est: {},
    direccion_est: {},
    telefono_est: {},
    fecha_nacimiento_est: {},
    telefono_est: {},
    edad_est: {},
    calificacion: {},
    nombre_est: {},
    apellido_est: {},
    correo_est: {},
  };

  const keys = Object.keys(memo);

  // Clasificar la informacion cargada en la variable data
  for (let i = 1; i < data.length; i++) {
    for (let j = 0; j < data[0].length; j++) {
      const currentData = data[i][j];

      // Verificar si el dato de una columna está vacío o es inválido(#REF!)
      if (!isValidString(currentData)) {
        iter(i, j, memo[keys[j]]);
        continue;
      }

      // Comprobar reglas de cada columna (cedula, email)
      if (!constraints(keys[j], currentData+"")) {
        iter(i, j, memo[keys[j]]);
      }
    }
  }

  const total = { ...memo };

  for (const key in total) {
    let count = 0;
    for (const data in total[key]) {
      count += parseInt(total[key][data]);
    }
    total[key] = count;
  }

  console.log(total);
};

const isValidString = (string) => {
  return !(string === null || string === "#ERROR_#REF!");
};

const constraints = (column, string) => {
  switch (column) {
    case "codigo_est":
      return isStudentCode(string);
    case "correo_est":
      return isEmail(string);
    case "cedula_est":
      return isIdentificationCard(string);
    case "direccion_est":
      return true;
    default:
      return true;
  }
};

const isStudentCode = (code) => {
  code = "" + code;
  return code.length === 11;
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

const iter = (row, index, memo) => {
  if (memo[row] !== undefined) {
  } else {
    memo[row] = 1;
  }
};

main();
