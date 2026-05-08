"use client";

import dynamic from "next/dynamic";

const InvitationDesigner = dynamic(() => import("./InvitationDesigner"), {
  ssr: false,
  loading: () => <div className="boot-screen">Loading editor...</div>
});

export default function Page() {
  return <InvitationDesigner />;
}
