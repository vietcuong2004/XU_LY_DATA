export async function api(path, data, { signal } = {}) {
  const options = data === undefined ? {} : {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)};
  const response = await fetch(path, {...options, signal});
  const result = await response.json();
  if (!response.ok) {
    const error = new Error(result.error || 'Không thể kết nối. Vui lòng thử lại.');
    error.status = response.status; error.code = result.code; throw error;
  }
  return result;
}
