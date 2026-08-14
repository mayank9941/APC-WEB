import ApcReport from "@/components/ApcReport";

export default async function PlazaPage({
  params,
}: {
  params: Promise<{ code: string }>;
}) {
  const { code } = await params;
  return (
    <main className="container">
      <ApcReport code={code} />
    </main>
  );
}
