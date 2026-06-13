import { DaybookPerson } from "../types/daybook";

export function AvatarStack({ people }: { people: DaybookPerson[] }) {
  const visible = people.slice(0, 3);
  if (!visible.length) return null;

  return (
    <div className="people-row">
      <div className="avatars">
        {visible.map((person, index) => {
          const label = person.displayName || person.email || "Person";
          return <div className="avatar" title={label} key={`${person.id || person.email || label}-${index}`}>{initials(label)}</div>;
        })}
        {people.length > visible.length && <div className="avatar overflow">+{people.length - visible.length}</div>}
      </div>
      <span>{people.map((person) => person.displayName || person.email).filter(Boolean).slice(0, 2).join(", ")}</span>
    </div>
  );
}

function initials(label: string): string {
  const parts = label.trim().split(/\s+/);
  return parts.slice(0, 2).map((part) => part[0]).join("").toUpperCase();
}
