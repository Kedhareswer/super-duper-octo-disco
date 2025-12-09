"use client";

import { useState, useEffect, useCallback } from "react";

const API_BASE = "http://127.0.0.1:8000";

type EvalScores = {
  preservation_score: number;
  instruction_adherence: number;
  fluency_score: number;
  overall_score: number;
};

type RecentEval = {
  original: string;
  edited: string;
  instruction: string;
  intent: string;
  confidence?: number;
  scores: EvalScores;
  timestamp: string;
};

type DashboardStats = {
  total_evals: number;
  avg_overall_score: number;
  avg_preservation: number;
  avg_adherence: number;
  avg_fluency: number;
  recent_evals: RecentEval[];
};

type TestResult = {
  original: string;
  instruction: string;
  edited: string;
  intent: string;
  confidence: number;
  eval_scores: EvalScores;
  timestamp: string;
};

type TestSuiteResult = {
  passed: number;
  total: number;
  pass_rate: number;
  results: Array<{
    name: string;
    passed: boolean;
    original: string;
    expected_contains: string;
    actual: string;
    scores: EvalScores;
  }>;
};

export default function EvalsPage() {
  const [stats, setStats] = useState<DashboardStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [testInput, setTestInput] = useState("");
  const [testInstruction, setTestInstruction] = useState("");
  const [testResult, setTestResult] = useState<TestResult | null>(null);
  const [testSuiteResult, setTestSuiteResult] = useState<TestSuiteResult | null>(null);
  const [running, setRunning] = useState(false);

  const fetchStats = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/evals/dashboard`);
      if (res.ok) {
        const data = await res.json();
        setStats(data);
      }
    } catch (e) {
      console.error("Failed to fetch stats:", e);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchStats();
    // Poll every 5 seconds
    const interval = setInterval(fetchStats, 5000);
    return () => clearInterval(interval);
  }, [fetchStats]);

  const runTest = async () => {
    if (!testInput.trim() || !testInstruction.trim()) return;
    setRunning(true);
    setTestResult(null);
    try {
      const res = await fetch(`${API_BASE}/evals/test`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          original: testInput,
          instruction: testInstruction,
        }),
      });
      if (res.ok) {
        const data = await res.json();
        setTestResult(data);
        fetchStats(); // Refresh stats
      }
    } catch (e) {
      console.error("Test failed:", e);
    } finally {
      setRunning(false);
    }
  };

  const runTestSuite = async () => {
    setRunning(true);
    setTestSuiteResult(null);
    try {
      const res = await fetch(`${API_BASE}/evals/test-suite`, {
        method: "POST",
      });
      if (res.ok) {
        const data = await res.json();
        setTestSuiteResult(data);
        fetchStats();
      }
    } catch (e) {
      console.error("Test suite failed:", e);
    } finally {
      setRunning(false);
    }
  };

  const clearHistory = async () => {
    try {
      await fetch(`${API_BASE}/evals/history`, { method: "DELETE" });
      fetchStats();
    } catch (e) {
      console.error("Failed to clear history:", e);
    }
  };

  const ScoreBar = ({ label, score, color }: { label: string; score: number; color: string }) => (
    <div className="mb-2">
      <div className="flex justify-between text-sm mb-1">
        <span>{label}</span>
        <span className="font-mono">{(score * 100).toFixed(1)}%</span>
      </div>
      <div className="h-2 bg-zinc-200 rounded-full overflow-hidden">
        <div
          className={`h-full ${color} transition-all duration-300`}
          style={{ width: `${score * 100}%` }}
        />
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-zinc-50 p-6">
      <div className="max-w-6xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <h1 className="text-2xl font-bold text-zinc-800">AI Edit Evaluation Dashboard</h1>
          <a href="/" className="text-blue-600 hover:underline text-sm">
            ← Back to Editor
          </a>
        </div>

        {/* Stats Cards */}
        <div className="grid grid-cols-1 md:grid-cols-5 gap-4 mb-6">
          <div className="bg-white rounded-lg shadow p-4">
            <div className="text-sm text-zinc-500">Total Evals</div>
            <div className="text-2xl font-bold text-zinc-800">
              {stats?.total_evals ?? 0}
            </div>
          </div>
          <div className="bg-white rounded-lg shadow p-4">
            <div className="text-sm text-zinc-500">Avg Overall</div>
            <div className="text-2xl font-bold text-emerald-600">
              {((stats?.avg_overall_score ?? 0) * 100).toFixed(1)}%
            </div>
          </div>
          <div className="bg-white rounded-lg shadow p-4">
            <div className="text-sm text-zinc-500">Preservation</div>
            <div className="text-2xl font-bold text-blue-600">
              {((stats?.avg_preservation ?? 0) * 100).toFixed(1)}%
            </div>
          </div>
          <div className="bg-white rounded-lg shadow p-4">
            <div className="text-sm text-zinc-500">Adherence</div>
            <div className="text-2xl font-bold text-purple-600">
              {((stats?.avg_adherence ?? 0) * 100).toFixed(1)}%
            </div>
          </div>
          <div className="bg-white rounded-lg shadow p-4">
            <div className="text-sm text-zinc-500">Fluency</div>
            <div className="text-2xl font-bold text-amber-600">
              {((stats?.avg_fluency ?? 0) * 100).toFixed(1)}%
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Test Panel */}
          <div className="bg-white rounded-lg shadow p-4">
            <h2 className="font-semibold text-zinc-800 mb-4">Run Test</h2>
            <div className="space-y-3">
              <div>
                <label className="block text-sm text-zinc-600 mb-1">Original Text</label>
                <textarea
                  className="w-full border rounded p-2 text-sm"
                  rows={3}
                  value={testInput}
                  onChange={(e) => setTestInput(e.target.value)}
                  placeholder="Enter text to edit..."
                />
              </div>
              <div>
                <label className="block text-sm text-zinc-600 mb-1">Instruction</label>
                <input
                  type="text"
                  className="w-full border rounded p-2 text-sm"
                  value={testInstruction}
                  onChange={(e) => setTestInstruction(e.target.value)}
                  placeholder="e.g., make it formal, fix grammar..."
                />
              </div>
              <div className="flex gap-2">
                <button
                  onClick={runTest}
                  disabled={running}
                  className="px-4 py-2 bg-blue-600 text-white rounded text-sm hover:bg-blue-700 disabled:opacity-50"
                >
                  {running ? "Running..." : "Run Test"}
                </button>
                <button
                  onClick={runTestSuite}
                  disabled={running}
                  className="px-4 py-2 bg-purple-600 text-white rounded text-sm hover:bg-purple-700 disabled:opacity-50"
                >
                  Run Test Suite
                </button>
              </div>
            </div>

            {/* Test Result */}
            {testResult && (
              <div className="mt-4 p-3 bg-zinc-50 rounded border">
                <div className="text-sm font-medium text-zinc-700 mb-2">Result</div>
                <div className="text-sm space-y-1">
                  <div><span className="text-zinc-500">Intent:</span> {testResult.intent}</div>
                  <div><span className="text-zinc-500">Confidence:</span> {(testResult.confidence * 100).toFixed(1)}%</div>
                  <div className="mt-2 p-2 bg-white rounded border">
                    <div className="text-xs text-zinc-500 mb-1">Edited Text:</div>
                    <div className="text-sm">{testResult.edited}</div>
                  </div>
                </div>
                <div className="mt-3">
                  <ScoreBar label="Overall" score={testResult.eval_scores.overall_score} color="bg-emerald-500" />
                  <ScoreBar label="Preservation" score={testResult.eval_scores.preservation_score} color="bg-blue-500" />
                  <ScoreBar label="Adherence" score={testResult.eval_scores.instruction_adherence} color="bg-purple-500" />
                  <ScoreBar label="Fluency" score={testResult.eval_scores.fluency_score} color="bg-amber-500" />
                </div>
              </div>
            )}

            {/* Test Suite Result */}
            {testSuiteResult && (
              <div className="mt-4 p-3 bg-zinc-50 rounded border">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-sm font-medium text-zinc-700">Test Suite Results</div>
                  <div className={`text-sm font-bold ${testSuiteResult.pass_rate >= 0.8 ? 'text-emerald-600' : 'text-red-600'}`}>
                    {testSuiteResult.passed}/{testSuiteResult.total} passed ({(testSuiteResult.pass_rate * 100).toFixed(0)}%)
                  </div>
                </div>
                <div className="space-y-2 max-h-60 overflow-auto">
                  {testSuiteResult.results.map((r, i) => (
                    <div key={i} className={`p-2 rounded text-xs ${r.passed ? 'bg-emerald-50 border-emerald-200' : 'bg-red-50 border-red-200'} border`}>
                      <div className="flex items-center gap-2">
                        <span>{r.passed ? '✓' : '✗'}</span>
                        <span className="font-medium">{r.name}</span>
                      </div>
                      {!r.passed && (
                        <div className="mt-1 text-zinc-600">
                          Expected "{r.expected_contains}" in "{r.actual}"
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Recent Evals */}
          <div className="bg-white rounded-lg shadow p-4">
            <div className="flex items-center justify-between mb-4">
              <h2 className="font-semibold text-zinc-800">Recent Evaluations</h2>
              <button
                onClick={clearHistory}
                className="text-xs text-red-600 hover:underline"
              >
                Clear History
              </button>
            </div>
            {loading ? (
              <div className="text-sm text-zinc-500">Loading...</div>
            ) : stats?.recent_evals.length === 0 ? (
              <div className="text-sm text-zinc-500">No evaluations yet. Run a test to get started.</div>
            ) : (
              <div className="space-y-3 max-h-[500px] overflow-auto">
                {stats?.recent_evals.map((eval_, i) => (
                  <div key={i} className="p-3 bg-zinc-50 rounded border text-sm">
                    <div className="flex items-start justify-between mb-2">
                      <div className="text-xs text-zinc-500">
                        {new Date(eval_.timestamp).toLocaleTimeString()}
                      </div>
                      <div className={`text-xs font-medium px-2 py-0.5 rounded ${
                        eval_.scores.overall_score >= 0.8 ? 'bg-emerald-100 text-emerald-700' :
                        eval_.scores.overall_score >= 0.5 ? 'bg-amber-100 text-amber-700' :
                        'bg-red-100 text-red-700'
                      }`}>
                        {(eval_.scores.overall_score * 100).toFixed(0)}%
                      </div>
                    </div>
                    <div className="text-xs text-zinc-600 mb-1">
                      <span className="font-medium">Instruction:</span> {eval_.instruction}
                    </div>
                    <div className="grid grid-cols-2 gap-2 text-xs">
                      <div className="p-1.5 bg-white rounded">
                        <div className="text-zinc-400 mb-0.5">Original</div>
                        <div className="truncate">{eval_.original}</div>
                      </div>
                      <div className="p-1.5 bg-white rounded">
                        <div className="text-zinc-400 mb-0.5">Edited</div>
                        <div className="truncate">{eval_.edited}</div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
