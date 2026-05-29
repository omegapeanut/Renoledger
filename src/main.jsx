import { Component, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import App from '../project-ledger-v2.jsx'

class ErrorBoundary extends Component {
  constructor(props) {
    super(props)
    this.state = { error: null }
  }
  static getDerivedStateFromError(error) {
    return { error }
  }
  render() {
    if (this.state.error) {
      return createElement('div', {style:{padding:32,fontFamily:'monospace',background:'#1a1a1a',color:'#f87171',minHeight:'100vh'}},
        createElement('div', {style:{fontSize:20,fontWeight:700,marginBottom:16}}, 'RenoLedger — Startup Error'),
        createElement('pre', {style:{whiteSpace:'pre-wrap',wordBreak:'break-all',fontSize:13}},
          this.state.error?.message + '\n\n' + this.state.error?.stack
        )
      )
    }
    return this.props.children
  }
}

createRoot(document.getElementById('root')).render(
  createElement(ErrorBoundary, null, createElement(App, null))
)
