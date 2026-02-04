// docs-app/src/App.tsx
import { BrowserRouter, Routes, Route } from 'react-router-dom';
import { Layout } from './components/Layout';
import { Home } from './pages/Home';
import { Vocabulary } from './pages/Vocabulary';
import { Clauses } from './pages/Clauses';
import { Controls } from './pages/Controls';
import { Merged } from './pages/Merged';
import { About } from './pages/About';
import './App.css';

function App() {
  return (
    <BrowserRouter basename="/iso27001-skos-taxonomy">
      <Routes>
        <Route element={<Layout />}>
          <Route index element={<Home />} />
          <Route path="/vocabulary" element={<Vocabulary />} />
          <Route path="/clauses" element={<Clauses />} />
          <Route path="/controls" element={<Controls />} />
          <Route path="/relationships" element={<Merged />} />
          <Route path="/about" element={<About />} />
        </Route>
      </Routes>
    </BrowserRouter>
  );
}

export default App;
