import React from 'react'
import ReactDOM from 'react-dom/client'
import App from '../project-ledger-v2.jsx'

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props)
    this.state = { error: null }
  }
  static getDerivedStateFromError(error) {
    return { error }
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{padding:32,fontFamily:'monospace',background:'#1a1a1a',color:'#f87171',minHeight:'100vh'}}>
          <div style={{fontSize:20,fontWeight:700,marginBottom:16}}>RenoLedger — Startup Error</div>
          <pre style={{whiteSpace:'pre-wrap',wordBreak:'break-all',fontSize:13}}>
            {this.state.error?.message}
            {'\n\n'}
            {this.state.error?.stack}
          </pre>
        </div>
      )
    }
    return this.props.children
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <ErrorBoundary>
    <App />
  </ErrorBoundary>
)
