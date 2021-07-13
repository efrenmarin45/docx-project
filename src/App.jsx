import React, { useState } from 'react'
import './App.css'
import generateDoc from './documentFile'

function App() {
  const [count, setCount] = useState(0)

  return (
    <div className="App">
      <header className="App-header">
        <h1>DOCX Project</h1>
        <p>
          <button type="button" onClick={generateDoc}>
            Generate Document
          </button>
        </p>
      </header>
    </div>
  )
}

export default App
