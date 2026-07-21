import "./globals.css";

export const metadata = {
  title: "Kartu Stok PFA",
  description: "Prototype Website",
  icons: {
    icon: "/icon.png",
  },
};

export default function RootLayout({
  children, 
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="id">
      <body>{children}</body>
    </html>
  );
}