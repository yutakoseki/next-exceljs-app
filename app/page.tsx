import React from 'react';
import UploadExcelComponent from './components/UploadExcelComponent';


const IndexPage: React.FC = () => {
  return (
    <div>
      <h1>Excel Processing Page</h1>
      <UploadExcelComponent />
      {/* 他のコンポーネントやUIを追加 */}
    </div>
  );
};

export default IndexPage;