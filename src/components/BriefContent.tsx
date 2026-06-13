export function BriefContent({ markdown }: { markdown: string }) {
  const nodes = [];
  const lines = markdown.split("\n");

  for (let index = 0; index < lines.length;) {
    const value = lines[index].trim();
    if (!value) {
      index += 1;
      continue;
    }
    if (value.startsWith("### ")) {
      nodes.push(<h4 key={`heading-${index}`}>{value.slice(4)}</h4>);
      index += 1;
      continue;
    }
    if (value.startsWith("## ") || value.startsWith("# ")) {
      nodes.push(<h3 key={`heading-${index}`}>{value.replace(/^#{1,2}\s+/, "")}</h3>);
      index += 1;
      continue;
    }
    if (/^[-*]\s+/.test(value)) {
      const entries = [];
      while (index < lines.length && /^[-*]\s+/.test(lines[index].trim())) {
        entries.push(lines[index].trim().replace(/^[-*]\s+/, ""));
        index += 1;
      }
      nodes.push(<ul key={`list-${index}`}>{entries.map((entry) => <li key={entry}>{entry}</li>)}</ul>);
      continue;
    }

    const paragraph = [];
    while (index < lines.length && lines[index].trim() && !/^(?:#{1,3}|[-*])\s+/.test(lines[index].trim())) {
      paragraph.push(lines[index].trim());
      index += 1;
    }
    nodes.push(<p key={`paragraph-${index}`}>{paragraph.join(" ")}</p>);
  }

  return <>{nodes}</>;
}
