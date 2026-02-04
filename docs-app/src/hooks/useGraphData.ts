import { useEffect, useState } from 'react';
import type { GraphData } from '../types';

export function useGraphData(dataFile: string) {
  const [data, setData] = useState<GraphData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    setError(null);

    fetch(`${import.meta.env.BASE_URL}data/${dataFile}`)
      .then(res => {
        if (!res.ok) throw new Error(`Failed to load ${dataFile}`);
        return res.json();
      })
      .then((json: GraphData) => {
        setData(json);
        setLoading(false);
      })
      .catch(err => {
        setError(err.message);
        setLoading(false);
      });
  }, [dataFile]);

  return { data, loading, error };
}
