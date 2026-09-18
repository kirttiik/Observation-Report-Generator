import { useState } from 'react'
import './App.css'

function App() {
  const [files, setFiles] = useState({ file1: null, file2: null, file3: null })
  const [config, setConfig] = useState({
    col_school: 1,
    col_teachers: 3,
    col_unique: 6,
    col_obs: 5,
    col_term_obs: 8
  })
  
  const [dates, setDates] = useState({
    year_start: '2025-04-01',
    year_end: '2026-04-01',
    t1_start: '2025-04-01',
    t1_end: '2025-10-15',
    t2_start: '2025-10-16',
    t2_end: '2026-04-01'
  })

  const [loading, setLoading] = useState(false)

  const handleFileChange = (e, fileKey) => {
    const file = e.target.files[0]
    setFiles(prev => ({ ...prev, [fileKey]: file }))
  }

  const handleConfigChange = (e) => {
    const { name, value } = e.target
    setConfig(prev => ({ ...prev, [name]: parseInt(value) || value }))
  }

  const handleDateChange = (e) => {
    const { name, value } = e.target
    setDates(prev => ({ ...prev, [name]: value }))
  }

  const generateReport = async () => {
    if (!files.file1) {
      alert("Please upload the Entire Year base report.")
      return
    }

    setLoading(true)
    const formData = new FormData()
    formData.append('file1', files.file1)
    if (files.file2) formData.append('file2', files.file2)
    if (files.file3) formData.append('file3', files.file3)
    
    Object.keys(config).forEach(key => {
      formData.append(key, config[key])
    })

    Object.keys(dates).forEach(key => {
      formData.append(key, dates[key])
    })

    try {
      const response = await fetch('/api/generate', {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        throw new Error('Failed to generate report.')
      }

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'Observation_Reports.zip'
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)
    } catch (error) {
      console.error(error)
      alert("Error generating report. Check terminal for details.")
    } finally {
      setLoading(false)
    }
  }

  return (
    <>
      <div className="header">
        <h1>Report Generator</h1>
        <p>Convert and combine your Observation Reports seamlessly via FastAPI</p>
      </div>

      <div className="container">
        {/* Cards Grid */}
        <div className="cards-grid">
          {/* Entire Year Card */}
          <div className="card">
            <div className="file-upload-wrapper">
              <input type="file" accept=".xls,.xlsx" onChange={(e) => handleFileChange(e, 'file1')} />
              <div className="card-header">
                <div className="file-icon">📄</div>
                <div className="card-title-group">
                  <h3>Entire Year</h3>
                  <span>Required Report</span>
                </div>
              </div>
              <div className="file-name" title={files.file1 ? files.file1.name : "Click to select base report"}>
                📁 {files.file1 ? files.file1.name : "Select Base Report"}
              </div>
            </div>
            <div className="date-range-group">
              <label>Date Range</label>
              <div className="date-inputs">
                <div className="date-input-wrapper">
                  <input type="date" name="year_start" value={dates.year_start} onChange={handleDateChange} />
                </div>
                <span className="to-text">to</span>
                <div className="date-input-wrapper">
                  <input type="date" name="year_end" value={dates.year_end} onChange={handleDateChange} />
                </div>
              </div>
            </div>
          </div>

          {/* Term 1 Card */}
          <div className="card">
            <div className="file-upload-wrapper">
              <input type="file" accept=".xls,.xlsx" onChange={(e) => handleFileChange(e, 'file2')} />
              <div className="card-header">
                <div className="file-icon">📄</div>
                <div className="card-title-group">
                  <h3>Term 1</h3>
                  <span>Optional Report</span>
                </div>
              </div>
              <div className="file-name" title={files.file2 ? files.file2.name : "Click to select Term 1 report"}>
                📁 {files.file2 ? files.file2.name : "Select Term 1 Report"}
              </div>
            </div>
            <div className="date-range-group">
              <label>Date Range</label>
              <div className="date-inputs">
                <div className="date-input-wrapper">
                  <input type="date" name="t1_start" value={dates.t1_start} onChange={handleDateChange} />
                </div>
                <span className="to-text">to</span>
                <div className="date-input-wrapper">
                  <input type="date" name="t1_end" value={dates.t1_end} onChange={handleDateChange} />
                </div>
              </div>
            </div>
          </div>

          {/* Term 2 Card */}
          <div className="card">
            <div className="file-upload-wrapper">
              <input type="file" accept=".xls,.xlsx" onChange={(e) => handleFileChange(e, 'file3')} />
              <div className="card-header">
                <div className="file-icon">📄</div>
                <div className="card-title-group">
                  <h3>Term 2</h3>
                  <span>Optional Report</span>
                </div>
              </div>
              <div className="file-name" title={files.file3 ? files.file3.name : "Click to select Term 2 report"}>
                📁 {files.file3 ? files.file3.name : "Select Term 2 Report"}
              </div>
            </div>
            <div className="date-range-group">
              <label>Date Range</label>
              <div className="date-inputs">
                <div className="date-input-wrapper">
                  <input type="date" name="t2_start" value={dates.t2_start} onChange={handleDateChange} />
                </div>
                <span className="to-text">to</span>
                <div className="date-input-wrapper">
                  <input type="date" name="t2_end" value={dates.t2_end} onChange={handleDateChange} />
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Column Mapping Section */}
        <div className="mapping-card">
          <div className="mapping-title">Column Mapping</div>
          <div className="mapping-subtitle">(1-based index)</div>
          
          <div className="mapping-grid">
            <div className="mapping-field">
              <label>School Name</label>
              <input type="number" name="col_school" value={config.col_school} onChange={handleConfigChange} />
            </div>
            <div className="mapping-field">
              <label>No. of Teachers</label>
              <input type="number" name="col_teachers" value={config.col_teachers} onChange={handleConfigChange} />
            </div>
            <div className="mapping-field">
              <label>Observation Till Date</label>
              <input type="number" name="col_obs" value={config.col_obs} onChange={handleConfigChange} />
            </div>
            <div className="mapping-field">
              <label>Unique Teachers</label>
              <input type="number" name="col_unique" value={config.col_unique} onChange={handleConfigChange} />
            </div>
            <div className="mapping-field">
              <label>Term 1 & 2 Obs</label>
              <input type="number" name="col_term_obs" value={config.col_term_obs} onChange={handleConfigChange} />
            </div>
          </div>
        </div>

        {/* Submit Button */}
        <button 
          className={`submit-btn ${loading ? 'loading' : ''}`}
          onClick={generateReport}
          disabled={loading || !files.file1}
        >
          {loading ? 'Processing Data...' : 'Generate Report (Excel + Images)'}
        </button>
      </div>
    </>
  )
}

export default App
