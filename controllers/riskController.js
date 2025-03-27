const sheetsService = require('../services/sheetsService');
const { ValidationError } = require('../middleware/errorHandler');

/**
 * Obtiene todos los riesgos asociados a una solicitud
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const getRiesgos = async (req, res) => {
  try {
    const { id_solicitud } = req.query;
    
    if (!id_solicitud) {
      return res.status(400).json({ 
        success: false,
        error: 'El ID de solicitud es requerido' 
      });
    }

    const riesgos = await sheetsService.getRisksBySolicitud(id_solicitud);
    
    res.status(200).json({
      success: true,
      data: riesgos
    });
  } catch (error) {
    console.error('Error al obtener riesgos:', error);
    res.status(500).json({
      success: false,
      error: 'Error al obtener riesgos de la solicitud',
      details: error.message
    });
  }
};

/**
 * Añade un nuevo riesgo para una solicitud
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const addRiesgo = async (req, res) => {
  try {
    const { nombre_riesgo, aplica, mitigacion, id_solicitud, categoria } = req.body;
    
    // Validación de datos
    if (!nombre_riesgo || !id_solicitud) {
      throw new ValidationError('nombre_riesgo y id_solicitud son campos requeridos');
    }

    // Obtener el último ID de riesgo
    const lastId = await sheetsService.getLastId('RIESGOS');
    const newId = lastId + 1;
    
    // Preparar el nuevo riesgo
    const riesgoData = {
      id_riesgo: newId.toString(),
      nombre_riesgo,
      aplica: aplica || 'No',
      mitigacion: mitigacion || '',
      id_solicitud: id_solicitud.toString(),
      categoria: categoria || 'General'
    };
    
    // Guardar el riesgo
    await sheetsService.saveRisk(riesgoData);
    
    res.status(201).json({
      success: true,
      message: 'Riesgo creado exitosamente',
      data: riesgoData
    });
  } catch (error) {
    console.error('Error al crear riesgo:', error);
    
    if (error instanceof ValidationError) {
      return res.status(400).json({
        success: false,
        error: error.message
      });
    }
    
    res.status(500).json({
      success: false,
      error: 'Error al crear riesgo',
      details: error.message
    });
  }
};

/**
 * Actualiza un riesgo existente
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const updateRiesgo = async (req, res) => {
  try {
    const { id_riesgo, nombre_riesgo, aplica, mitigacion, categoria } = req.body;
    
    if (!id_riesgo) {
      throw new ValidationError('ID de riesgo es requerido');
    }
    
    // Verificar si el riesgo existe
    const risksData = await sheetsService.getRiskById(id_riesgo);
    
    if (!risksData) {
      return res.status(404).json({
        success: false,
        error: `No se encontró riesgo con ID ${id_riesgo}`
      });
    }
    
    // Actualizar los campos modificados
    const updatedData = {
      ...risksData,
      nombre_riesgo: nombre_riesgo || risksData.nombre_riesgo,
      aplica: aplica !== undefined ? aplica : risksData.aplica,
      mitigacion: mitigacion !== undefined ? mitigacion : risksData.mitigacion,
      categoria: categoria || risksData.categoria
    };
    
    // Actualizar el riesgo
    await sheetsService.updateRisk(updatedData);
    
    res.status(200).json({
      success: true,
      message: 'Riesgo actualizado exitosamente',
      data: updatedData
    });
  } catch (error) {
    console.error('Error al actualizar riesgo:', error);
    
    if (error instanceof ValidationError) {
      return res.status(400).json({
        success: false,
        error: error.message
      });
    }
    
    res.status(500).json({
      success: false,
      error: 'Error al actualizar riesgo',
      details: error.message
    });
  }
};

/**
 * Elimina un riesgo existente
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const deleteRiesgo = async (req, res) => {
  try {
    const { id_riesgo } = req.params;
    
    if (!id_riesgo) {
      throw new ValidationError('ID de riesgo es requerido');
    }
    
    // Verificar si el riesgo existe
    const riskExists = await sheetsService.getRiskById(id_riesgo);
    
    if (!riskExists) {
      return res.status(404).json({
        success: false,
        error: `No se encontró riesgo con ID ${id_riesgo}`
      });
    }
    
    // Eliminar el riesgo
    await sheetsService.deleteRisk(id_riesgo);
    
    res.status(200).json({
      success: true,
      message: 'Riesgo eliminado exitosamente'
    });
  } catch (error) {
    console.error('Error al eliminar riesgo:', error);
    
    if (error instanceof ValidationError) {
      return res.status(400).json({
        success: false,
        error: error.message
      });
    }
    
    res.status(500).json({
      success: false,
      error: 'Error al eliminar riesgo',
      details: error.message
    });
  }
};

/**
 * Obtiene categorías predefinidas de riesgos
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const getCategoriasRiesgo = async (req, res) => {
  try {
    // Categorías predefinidas de riesgos
    const categorias = [
      { id: 'diseno', nombre: 'Diseño', descripcion: 'Riesgos relacionados con la fase de diseño' },
      { id: 'locacion', nombre: 'Locación', descripcion: 'Riesgos relacionados con la ubicación' },
      { id: 'desarrollo', nombre: 'Desarrollo', descripcion: 'Riesgos durante la fase de desarrollo' },
      { id: 'cierre', nombre: 'Cierre', descripcion: 'Riesgos en la etapa de cierre' },
      { id: 'otros', nombre: 'Otros', descripcion: 'Otros tipos de riesgos' }
    ];
    
    res.status(200).json({
      success: true,
      data: categorias
    });
  } catch (error) {
    console.error('Error al obtener categorías:', error);
    res.status(500).json({
      success: false,
      error: 'Error al obtener categorías de riesgos',
      details: error.message
    });
  }
};

/**
 * Migra riesgos del formulario 3 al formato dinámico
 * @param {Object} req - Objeto de solicitud Express
 * @param {Object} res - Objeto de respuesta Express
 */
const migrarRiesgosForm3 = async (req, res) => {
  try {
    const { id_solicitud } = req.body;
    
    if (!id_solicitud) {
      throw new ValidationError('ID de solicitud es requerido');
    }
    
    // Obtener datos del formulario 3
    const hojas = {
      SOLICITUDES3: {
        range: 'SOLICITUDES3!A2:AC',
        fields: sheetsService.fieldDefinitions.SOLICITUDES3
      }
    };
    
    const solicitudData = await sheetsService.getSolicitudData(id_solicitud, hojas);
    
    if (!solicitudData.SOLICITUDES3) {
      return res.status(404).json({
        success: false,
        error: `No se encontraron datos del formulario 3 para la solicitud ${id_solicitud}`
      });
    }
    
    const form3Data = solicitudData.SOLICITUDES3;
    
    // Mapeo de riesgos del formulario 3
    const riesgosMap = [
      { field: 'aplicaDiseno1', nombre: 'Riesgo de diseño 1', categoria: 'diseno' },
      { field: 'aplicaDiseno2', nombre: 'Riesgo de diseño 2', categoria: 'diseno' },
      { field: 'aplicaDiseno3', nombre: 'Riesgo de diseño 3', categoria: 'diseno' },
      { field: 'aplicaLocacion1', nombre: 'Riesgo de locación 1', categoria: 'locacion' },
      { field: 'aplicaLocacion2', nombre: 'Riesgo de locación 2', categoria: 'locacion' },
      { field: 'aplicaLocacion3', nombre: 'Riesgo de locación 3', categoria: 'locacion' },
      { field: 'aplicaDesarrollo1', nombre: 'Riesgo de desarrollo 1', categoria: 'desarrollo' },
      { field: 'aplicaDesarrollo2', nombre: 'Riesgo de desarrollo 2', categoria: 'desarrollo' },
      { field: 'aplicaDesarrollo3', nombre: 'Riesgo de desarrollo 3', categoria: 'desarrollo' },
      { field: 'aplicaDesarrollo4', nombre: 'Riesgo de desarrollo 4', categoria: 'desarrollo' },
      { field: 'aplicaCierre1', nombre: 'Riesgo de cierre 1', categoria: 'cierre' },
      { field: 'aplicaCierre2', nombre: 'Riesgo de cierre 2', categoria: 'cierre' },
      { field: 'aplicaOtros1', nombre: 'Otro riesgo 1', categoria: 'otros' },
      { field: 'aplicaOtros2', nombre: 'Otro riesgo 2', categoria: 'otros' }
    ];
    
    // Obtener último ID de riesgo para asignar nuevos IDs
    let lastId = await sheetsService.getLastId('RIESGOS');
    
    // Convertir riesgos del formulario 3 al nuevo formato
    const nuevosRiesgos = [];
    for (const riesgo of riesgosMap) {
      if (form3Data[riesgo.field] === 'Sí' || form3Data[riesgo.field] === 'Si' || form3Data[riesgo.field] === true) {
        lastId++;
        nuevosRiesgos.push({
          id_riesgo: lastId.toString(),
          nombre_riesgo: riesgo.nombre,
          aplica: 'Sí',
          mitigacion: '',
          id_solicitud: id_solicitud.toString(),
          categoria: riesgo.categoria
        });
      }
    }
    
    // Guardar los nuevos riesgos
    if (nuevosRiesgos.length > 0) {
      await sheetsService.saveBulkRisks(nuevosRiesgos);
    }
    
    res.status(200).json({
      success: true,
      message: `Se migraron ${nuevosRiesgos.length} riesgos del formulario 3`,
      data: nuevosRiesgos
    });
  } catch (error) {
    console.error('Error al migrar riesgos:', error);
    
    if (error instanceof ValidationError) {
      return res.status(400).json({
        success: false,
        error: error.message
      });
    }
    
    res.status(500).json({
      success: false,
      error: 'Error al migrar riesgos del formulario 3',
      details: error.message
    });
  }
};

module.exports = {
  getRiesgos,
  addRiesgo,
  updateRiesgo,
  deleteRiesgo,
  getCategoriasRiesgo,
  migrarRiesgosForm3
};