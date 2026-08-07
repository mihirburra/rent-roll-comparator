export const metadata = {
  title: "Property Status Report Dashboard",
  description: "Monthly status report comparator and decision engine",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body style={{ margin: 0 }}>{children}</body>
    </html>
  );
}
