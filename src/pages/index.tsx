import { useRouter } from 'next/router';

const Home = () => {
  const router = useRouter();

  const handleClick = () => {
    router.push('/tool/Excel');
  };

  return (
    <div>
      <h1>This is Kyungho & Sunghee 's Page</h1>
      <h3 onClick={handleClick} style={{ cursor: 'pointer' }}>
        Click here to use our tool
      </h3>
    </div>
  );
};

export default Home;
