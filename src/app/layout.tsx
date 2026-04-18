import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Office Forge',
  description: 'A responsive document workspace for DOCX, PDF, PPT, and XLSX editing.'
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
