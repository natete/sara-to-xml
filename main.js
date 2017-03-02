const csvToJson = require('csvtojson');
const js2xmlparser = require('js2xmlparser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const args = require('yargs')
    .option('inputPath', {
      alias: 'ip',
      describe: 'the folder where your data source is'
    })
    .option('inputFilename', {
      alias: 'if',
      describe: 'name of the source file'
    })
    .option('ouputPath', {
      alias: 'op',
      describe: 'the folder to save your file'
    })
    .option('outputFilename', {
      alias: 'of',
      describe: 'name of the results file'
    })
    .option('year', {
      alias: 'y',
      describe: 'the year of the contracts to be parsed'
    })
    .help()
    .argv;

const data = {
  contratos: [],
  adjudicatarios: [],
  utes: [],
  presupuestarias: []
};

const done = {
  contratos: false,
  adjudicatarios: false,
  utes: false,
  presupuestarias: false
};

init();


function init() {
  console.log('Use --help to see the available options');

  const inputPath = args.inputPath || args.ip || './data';
  const inputFilename = (args.inputFilename || args.if || 'TCU_PRUEBA_SARA_CON LOTES') + '.xlsx';

  console.log(inputFilename);
  const workbook = xlsx.readFile(path.join(inputPath, inputFilename));

  workbook.SheetNames
      .forEach(sheetName => {
        const lowerSheetName = sheetName.toLowerCase();

        csvToJson({ delimiter: ';' })
            .fromString(xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName], { FS: ';' }))
            .on('json', (jsonObj) => data[lowerSheetName].push(jsonObj))
            .on('done', (error) => handleDoneParsing(error, lowerSheetName));
      });
}

function handleDoneParsing(error, key) {
  if (error) {
    console.error(error);
    process.exit(1);
  } else {

    done[key] = true;
    console.log(`Parsing ${key} finished`);

    if (Object.keys(done).every(key => done[key])) {
      console.log('Done parsing!!');
      processData();
    }
  }
}

function processData() {
  const contratos = data.contratos.map(contrato => processContrato(contrato));

  const outputPath = args.outputPath || args.op || './result';
  const filename = (args.outputFilename || args.of || 'result') + '.xml';
  const year = args.year || args.y || 2016;

  if (!fs.existsSync(outputPath)) {
    fs.mkdirSync(outputPath);
  }

  const resultPath = path.join(outputPath, filename);

  const result = js2xmlparser.parse('Rendicion',
      {
        '@': { ejercicio: year },
        Contrato: contratos
      },
      {
        wrapHandlers: {
          Adjudicatarios: () => 'Adjudicatario',
          Entidades: () => 'Entidad',
          AplicacionesPresupuestarias: () => 'AplicacionPresupuestaria'
        },
        format: {
          doubleQuotes: true
        }
      }
  );

  fs.writeFileSync(resultPath, result);

  console.log(`Success! You can open your file in ${resultPath}`)
}

function processContrato(contrato) {

  const result = {
    RefContrato: contrato.RefContrato,
    FechaAdjudicacion: contrato.FechaAdjudicacion,
    FechaFormalizacion: contrato.FechaFormalizacion,
    DescTipoContrato: contrato.DescTipoContrato,
    DescFormaTramitacion: contrato.DescFormaTramitacion,
    DescLegislacionAplicable: contrato.DescLegislacionAplicable,
    DescProcAdjudicacion: contrato.DescProcAdjudicacion,
    Sara: contrato.Sara,
    ValorEstimado: typeof contrato.ValorEstimado === 'number' ? contrato.ValorEstimado.toLocaleString() : contrato.ValorEstimado,
    NumLotes: contrato.NumLotes,
    Objeto: contrato.Objeto,
    ImporteAdjudicacion: typeof contrato.ImporteAdjudicacion === 'number' ? contrato.ImporteAdjudicacion.toLocaleString() : contrato.ImporteAdjudicacion,
    Impuestos: typeof contrato.Impuestos === 'number' ? contrato.Impuestos.toLocaleString() : contrato.Impuestos,
    PresupuestoLicitacion: typeof contrato.PresupuestoLicitacion === 'number' ? contrato.PresupuestoLicitacion.toLocaleString() : contrato.PresupuestoLicitacion,
    PlazoEjecucionMeses: contrato.PlazoEjecucionMeses
  };

  result.Adjudicatarios = getAdjudicatarios(contrato.RefContrato);

  result.AplicacionesPresupuestarias = getAplicacionesPresupuestarias(contrato.RefContrato);

  return result;
}

function getAdjudicatarios(refContrato) {
  return data.adjudicatarios
      .filter(adjudicatario => adjudicatario.RefContrato === refContrato)
      .map(adjudicatario => getAdjudicatario(adjudicatario));
}


function getAdjudicatario(adjudicatario) {
  const result = {
    Extranjero: adjudicatario.Extranjero,
    Cif: adjudicatario.Cif,
    Nombre: adjudicatario.Nombre,
    Ute: adjudicatario.Ute
  };

  if (adjudicatario.Ute) {
    result.Entidades = getEntidades(adjudicatario.Cif);
  }

  return result;
}

function getEntidades(cif) {
  return data.utes
      .filter(entidad => entidad['CIF Adjudicatario'] === cif)
      .map(entidad => getEntidad(entidad));
}

function getEntidad(entidad) {
  return {
    Extranjero: entidad.Extranjero,
    Cif: entidad.Cif,
    Nombre: entidad.Nombre
  };
}

function getAplicacionesPresupuestarias(refContrato) {
  return data.presupuestarias
      .filter(presupuestaria => presupuestaria.RefContrato === refContrato)
      .map(presupuestaria => getAplicacionPresupuestaria(presupuestaria));
}

function getAplicacionPresupuestaria(presupuestaria) {
  return {
    Descripcion: presupuestaria.Descripcion,
    Importe: typeof presupuestaria.Importe === 'number' ? presupuestaria.Importe.toLocaleString() : presupuestaria.Importe
  };
}