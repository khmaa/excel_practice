const Layout = ({ children }: { children: React.ReactNode }) => {
  return (
    <div className="container">
      <main>{children}</main>
    </div>
  );
};

export default Layout;
