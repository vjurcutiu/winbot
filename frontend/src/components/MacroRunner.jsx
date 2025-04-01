import React, { useState } from 'react';

const MacroRunner = () => {
  const [file, setFile] = useState(null);
  const [macroName, setMacroName] = useState('');
  const [response, setResponse] = useState(null);
  const [error, setError] = useState(null);

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleMacroNameChange = (e) => {
    setMacroName(e.target.value);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file || !macroName) {
      setError('Please provide both a file and a macro name.');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);
    formData.append('macro_name', macroName);

    try {
      const res = await fetch('/run_macro', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (res.ok) {
        setResponse(data);
        setError(null);
      } else {
        setError(data.message || 'Error occurred while running the macro.');
        setResponse(null);
      }
    } catch (err) {
      setError('An error occurred while uploading the file.');
      setResponse(null);
    }
  };

  return (
    <div>
      <h2>Run Macro on File</h2>
      <form onSubmit={handleSubmit}>
        <div>
          <label htmlFor="file">Choose file:</label>
          <input type="file" id="file" onChange={handleFileChange} />
        </div>
        <div>
          <label htmlFor="macroName">Macro Name:</label>
          <input
            type="text"
            id="macroName"
            value={macroName}
            onChange={handleMacroNameChange}
          />
        </div>
        <button type="submit">Run Macro</button>
      </form>
      {response && (
        <div>
          <h3>Success:</h3>
          <p>{response.log}</p>
          <a href={response.output_file} target="_blank" rel="noopener noreferrer">
            Download Processed File
          </a>
        </div>
      )}
      {error && <p style={{ color: 'red' }}>Error: {error}</p>}
    </div>
  );
};

export default MacroRunner;
