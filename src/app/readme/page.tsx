import Link from 'next/link';

export default function ReadmePage() {
  return (
    <main className="min-h-screen p-4 sm:p-6 lg:p-8">
      <div className="mx-auto max-w-5xl rounded-[28px] border border-white/10 bg-[#08101e]/90 p-6 text-white shadow-soft backdrop-blur-xl sm:p-8">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <h1 className="text-2xl font-semibold sm:text-3xl">Project Readme</h1>
          <Link href="/" className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm transition hover:bg-white/10">
            Back to Editor
          </Link>
        </div>

        <section className="mt-6 space-y-4 text-white/85">
          <p>
            Many people need to edit forms, resumes, school files, and business documents, but do not have a Microsoft Office subscription.
            A lot of online alternatives also introduce payment walls, limits, or locked features.
          </p>
          <p>
            Office Forge was built as a free and open source alternative so people can still work with Word (DOCX), spreadsheets (XLSX/CSV), presentations (PPT), and PDFs directly in the browser.
          </p>
          <p>
            The mission is practical access: if you can open a browser, you should still be able to get real document work done.
          </p>
        </section>

        <section className="mt-8 rounded-2xl border border-white/10 bg-white/5 p-4 sm:p-5">
          <h2 className="text-lg font-semibold text-white">Community Project</h2>
          <p className="mt-3 text-sm text-white/80">
            Developers are encouraged to join and improve this project together. Contributions are welcome for parser fidelity, editing UX, accessibility, mobile support, and performance.
          </p>
          <p className="mt-3 text-sm text-white/80">
            Contribute on GitHub:
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
        </section>
      </div>
    </main>
  );
}
