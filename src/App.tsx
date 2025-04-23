import './App.css'
import ExcelProcessor from './components/ExcelProcessor'

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>Excel链接自动打开工具</h1>
      </header>
      <main>
        <ExcelProcessor />
      </main>
      <footer>
        <p>© {new Date().getFullYear()} Excel链接处理工具</p>
      </footer>
    </div>
  )
}

export default App
