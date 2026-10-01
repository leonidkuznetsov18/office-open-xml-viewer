/** Exact squared Euclidean distance transform of a binary raster, using the
 * lower envelope of parabolas (Felzenszwalb/Huttenlocher). Two separable passes
 * take O(width·height) time and bounded raster-sized storage. The caller must
 * include a transparent border so an all-filled outline still has a boundary.
 * Distances are in raster pixels, independent of gradient colour or stops. */
export function interiorDistances(alpha: Uint8ClampedArray, width: number, height: number): Float32Array {
  const distances = new Float32Array(width * height);
  const limit = width * width + height * height + 1;
  for (let i = 0; i < distances.length; i++) {
    // Raster policy: coverage at the pixel center chooses the silhouette.
    distances[i] = alpha[i * 4 + 3] >= 128 ? limit : 0;
  }
  const size = Math.max(width, height);
  const input = new Float64Array(size);
  const output = new Float64Array(size);
  const vertices = new Int32Array(size);
  const crossings = new Float64Array(size + 1);
  const transform = (n: number) => {
    let k = 0;
    vertices[0] = 0; crossings[0] = -Infinity; crossings[1] = Infinity;
    for (let q = 1; q < n; q++) {
      let p = vertices[k];
      let s = ((input[q] + q * q) - (input[p] + p * p)) / (2 * (q - p));
      while (s <= crossings[k]) {
        p = vertices[--k];
        s = ((input[q] + q * q) - (input[p] + p * p)) / (2 * (q - p));
      }
      vertices[++k] = q; crossings[k] = s; crossings[k + 1] = Infinity;
    }
    k = 0;
    for (let q = 0; q < n; q++) {
      while (crossings[k + 1] < q) k++;
      const delta = q - vertices[k];
      output[q] = delta * delta + input[vertices[k]];
    }
  };
  for (let x = 0; x < width; x++) {
    for (let y = 0; y < height; y++) input[y] = distances[y * width + x];
    transform(height);
    for (let y = 0; y < height; y++) distances[y * width + x] = output[y];
  }
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) input[x] = distances[y * width + x];
    transform(width);
    // Background samples locate a pixel cell's center. Its boundary lies half
    // a texel nearer; convert center distance to raster-boundary clearance.
    for (let x = 0; x < width; x++) distances[y * width + x] = Math.max(0, Math.sqrt(output[x]) - .5);
  }
  return distances;
}
