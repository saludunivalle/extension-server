// Agregar a utils/apiUtils.js
async function withExponentialBackoff(operation, maxRetries = 5) {
  let retries = 0;
  while (retries < maxRetries) {
    try {
      return await operation();
    } catch (error) {
      if (error.code === 429 || error.message?.includes('Quota exceeded')) {
        const delay = Math.pow(2, retries) * 1000 + Math.random() * 1000;
        console.log(`Quota exceeded, retrying in ${delay}ms (retry ${retries+1}/${maxRetries})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        retries++;
      } else {
        throw error;
      }
    }
  }
  throw new Error(`Failed after ${maxRetries} retries`);
}

// Exportar la función
module.exports = {
  withExponentialBackoff
};