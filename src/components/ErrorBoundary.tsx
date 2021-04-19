import React from 'react'

interface ErrorBoundaryState {
  hasError: Boolean
}

export class ErrorBoundary extends React.Component<{}, ErrorBoundaryState> {
  constructor(props: any) {
    super(props)
    this.state = { hasError: false }
  }

  static getDerivedStateFromError(error: any) {
    // update state so the next render will show the fallback UI.
    return { hasError: true }
  }

  componentDidCatch(error: any, errorInfo: any) {
    console.error(error, errorInfo)
  }

  render() {
    if (this.state.hasError) {
      return (
        <p>
          Error: Something went wrong when trying to load this widget.
          <br />
          Please reload the page.
        </p>
      )
    }

    return this.props.children
  }
}
