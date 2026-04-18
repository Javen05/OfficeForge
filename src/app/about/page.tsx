import Link from 'next/link';

export default function AboutPage() {
  return (
    <main className="min-h-screen p-4 sm:p-6 lg:p-8">
      <div className="mx-auto max-w-4xl rounded-[28px] border border-white/10 bg-[#08101e]/90 p-6 text-white shadow-soft backdrop-blur-xl sm:p-8">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <h1 className="text-2xl font-semibold sm:text-3xl">About Office Forge</h1>
          <Link href="/" className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm transition hover:bg-white/10">
            Back to Editor
          </Link>
        </div>

        <p className="mt-6 text-white/85">
          Office Forge is an open source, browser-based workspace for DOCX, PDF, PPT, and XLSX files.
          It is built for people who need productivity tools on any device, including mobile, without being locked into expensive software.
        </p>

        <p className="mt-4 text-white/75">
          Repository:
          {' '}
          <a
            href="https://github.com/Javen05/OfficeForge"
            target="_blank"
            rel="noreferrer"
            className="text-[#f6c76a] underline decoration-[#f6c76a]/60 underline-offset-4"
          >
            github.com/Javen05/OfficeForge
          </a>
        </p>
      </div>
    </main>
  );
}
