// renderer/pages/home.jsx
import { useRef, useState } from 'react'
import Head from 'next/head'

function Home() {
  const [department, setDepartment] = useState('')
  const [isSubmitting, setIsSubmitting] = useState(false)
  const ref = useRef(null)
  const [uploads, setUploads] = useState({
    upload1: { files: null, visible: true },
    upload2: { files: null, visible: true },
    upload3: { files: null, visible: true },
  })

  const baseUrl = `https://arcchitect-nextjs-api.vercel.app`

  const handleFileChange = (id, files) => {
    setUploads(prev => ({
      ...prev,
      [id]: {
        files: files ? Array.from(files) : null,
        visible: false,
      }
    }))
  }

  const resetUpload = (id) => {
    setUploads(prev => ({
      ...prev,
      [id]: {
        files: null,
        visible: true,
      }
    }))
  }

  const handleSubmit = async (e) => {
    e.preventDefault()
    const allFilesUploaded = Object.values(uploads).every(upload => upload.files?.length)
    
    if (!allFilesUploaded) {
      ipcRenderer.send('show-alert', 'Please upload all required files before exporting.')
      return
    }

    setIsSubmitting(true)

    try {
      const formData = new FormData()
      
      Object.entries(uploads).forEach(([key, value]) => {
        if (key === 'upload2') {
          value.files?.forEach(file => {
            formData.append('file2', file)
          })
        } else {
          value.files?.[0] && formData.append(
            `file${key.replace('upload', '')}`,
            value.files[0]
          )
        }
      })

      const endpoint = department === '1' 
        ? `https://architect-excel.onrender.com/api/excel`
        : `${baseUrl}/api/excel/merge`

      const response = await fetch(endpoint, {
        method: 'POST',
        body: formData,
      })

      if (response.ok) {
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Final-output.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        alert('Downloaded Successfully');
    } else {
        throw new Error('Failed to download');
    }
    } catch (error) {
      console.error('Error:', error);
      alert('File upload failed');
    } finally {
      setIsSubmitting(false)
    }
  }

 

  const FileUpload = ({ id, label, sublabel, multiple = false, onFileChange }) => (
    <div className="file-upload">
      <label className="block text-sm font-medium text-gray-700 mb-2">
        {label}
        {sublabel && (
          <span className="text-[11px]"> {sublabel}</span>
        )}
      </label>
      <label className="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-gray-300 border-dashed rounded-lg cursor-pointer hover:border-blue-500 transition duration-150 ease-in-out" htmlFor={id}>
        <div className="space-y-1 text-center">
          <i className="fas fa-file-excel text-4xl text-gray-400"></i>
          <div className="flex text-sm text-gray-600">
            <span className="relative cursor-pointer bg-white rounded-md font-medium text-blue-600 hover:text-blue-500 focus-within:outline-none focus-within:ring-2 focus-within:ring-offset-2 focus-within:ring-blue-500">
              <span>Upload a file</span>
              <input
                id={id}
                type="file"
                className="sr-only"
                accept=".xlsx, .xls"
                multiple={multiple}
                onChange={(e) => onFileChange(e.target.files)}
                ref={ref}
              />
            </span>
            <p className="pl-1">or drag and drop</p>
          </div>
          <p className="text-xs text-gray-500">XLSX up to 10MB</p>
        </div>
      </label>
    </div>
  )

  return (
    <>
      <Head>
        <title>Drawing Excel Export</title>
        <link
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css"
          rel="stylesheet"
        />
      </Head>

      <div className="bg-gray-50 min-h-screen">
        <header className="bg-white shadow-sm">
          <nav className="container mx-auto px-6 py-4">
            <div className="flex justify-between items-center">
              <div className="text-2xl font-bold text-gray-800">ALA</div>
              <div className="space-x-6">
                <a
                  href="#"
                  className="text-gray-600 hover:text-gray-900 transition duration-150 ease-in-out"
                >
                  Home
                </a>
                <a
                  href="#"
                  className="text-gray-600 hover:text-gray-900 transition duration-150 ease-in-out"
                >
                  About
                </a>
                <a
                  href="#"
                  className="text-gray-600 hover:text-gray-900 transition duration-150 ease-in-out"
                >
                  Contact
                </a>
              </div>
            </div>
          </nav>
        </header>

        <main className="container mx-auto px-6 py-12">
          <div className="bg-white rounded-xl shadow-lg p-8 max-w-4xl mx-auto">
            <h1 className="text-3xl font-bold text-gray-900 mb-8 text-center">
              Drawing Excel export
            </h1>

            <form onSubmit={handleSubmit} className="space-y-8">
              <div>
                <label htmlFor="department" className="block text-sm font-medium text-gray-700 mb-2">
                  Select Department
                </label>
                <select
                  id="department"
                  value={department}
                  onChange={(e) => setDepartment(e.target.value)}
                  className="block w-full px-4 py-3 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 transition duration-150 ease-in-out"
                  required
                >
                  <option value="" disabled>Choose the department</option>
                  <option value="1">P&ID</option>
                  <option value="2">SLD</option>
                </select>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                {Object.entries(uploads).map(([id, upload]) => (
                  <div key={id}>
                    {upload.visible ? (
                      <FileUpload
                        id={id}
                        label={
                          id === 'upload1' ? 'EPE Excel File' :
                          id === 'upload2' ? 'MER Excel File' :
                          'SAP Excel File'
                        }
                        sublabel={id === 'upload2' ? '(Eg:-FAHN-6-50-0001-001)' : undefined}
                        multiple={id === 'upload2'}
                        onFileChange={(files) => handleFileChange(id, files)}
                      />
                    ) : (
                      <div className="bg-gray-100 rounded-lg p-4 flex items-center justify-between">
                        <span className="text-sm text-gray-700 truncate">
                          {upload.files?.map(f => f.name).join(', ')}
                        </span>
                        <button
                          type="button"
                          onClick={() => resetUpload(id)}
                          className="text-red-500 hover:text-red-700"
                        >
                          <i className="fas fa-times"></i>
                        </button>
                      </div>
                    )}
                  </div>
                ))}
              </div>

              <div className="flex justify-center">
                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="inline-flex items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition duration-150 ease-in-out disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isSubmitting ? (
                    <div className="flex items-center justify-center">
                      <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-white-500"></div>
                      <span className="ml-3 text-white-500 font-semibold">Exporting...</span>
                    </div>
                  ) : (
                    <>
                      <i className="fas fa-file-export mr-2"></i>
                      Export
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </main>
      </div>
    </>
  )
}

export default Home