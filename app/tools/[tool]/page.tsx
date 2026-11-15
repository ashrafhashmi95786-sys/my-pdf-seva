import ToolWorkspace from "../../tool-workspace";

const toolBySlug = {
  "merge-pdf": "merge",
  "split-pdf": "split",
  "remove-pages": "remove",
  "rotate-pdf": "rotate",
  "jpg-to-pdf": "jpg",
  "word-to-pdf": "word",
  "excel-to-pdf": "excel",
  "pdf-to-word": "pdfword",
  "pdf-to-excel": "pdfexcel"
} as const;

type ToolSlug = keyof typeof toolBySlug;

export function generateStaticParams() {
  return Object.keys(toolBySlug).map(tool => ({ tool }));
}

export default async function ToolPage({ params }: { params: Promise<{ tool: string }> }) {
  const { tool } = await params;
  if (!(tool in toolBySlug)) return null;
  return <ToolWorkspace initialTool={toolBySlug[tool as ToolSlug]} />;
}
