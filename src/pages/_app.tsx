// // pages/_app.tsx
// import { AppProps } from "next/app";

// function MyApp({ Component, pageProps }: AppProps) {
//   return <Component {...pageProps} />;
// }

// export default MyApp;
// src/pages/_app.tsx
import { AppProps } from 'next/app';
// import "../styles/global.css"; // 전역 스타일 가져오기
import Layout from '../components/Layout'; // 레이아웃 컴포넌트 가져오기

function MyApp({ Component, pageProps }: AppProps) {
  return (
    <Layout>
      <Component {...pageProps} />
    </Layout>
  );
}

export default MyApp;
