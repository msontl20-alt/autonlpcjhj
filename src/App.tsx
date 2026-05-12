import React, { useState, useEffect, useRef } from 'react';
import { 
  Upload, 
  FileDown, 
  Settings, 
  Users, 
  CheckCircle2, 
  AlertCircle, 
  Info, 
  Pencil, 
  X, 
  Save,
  ChevronRight,
  Database,
  LayoutDashboard
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';

// --- Constants & Types ---

const defaultCommentBank = `Năng lực/Phẩm chất,Mức,Nhận xét
Yêu nước,Tốt,"Em yêu thiên nhiên, quê hương và có những việc làm thiết thực bảo vệ thiên nhiên."
Yêu nước,Tốt,"Em tích cực tham gia các hoạt động tập thể và thực hiện tốt nội quy trường lớp."
Yêu nước,Tốt,"Em biết trân trọng, bảo vệ tự nhiên và tự hào về truyền thống quê hương."
Yêu nước,Tốt,"Em hiểu và tự hào về lịch sử, truyền thống cách mạng của quê hương, đất nước."
Yêu nước,Tốt,"Em luôn có ý thức giữ gìn và phát huy các giá trị văn hóa tốt đẹp của dân tộc."
Yêu nước,Đạt,"Em có ý thức thực hiện nội quy và tham gia các hoạt động chung."
Yêu nước,Đạt,"Em biết yêu quê hương đất nước qua các bài học."
Yêu nước,Đạt,"Em tham gia đầy đủ các hoạt động của trường lớp."
Yêu nước,Đạt,"Em có thái độ tôn trọng khi nghe kể chuyện về lịch sử dân tộc."
Yêu nước,Cần cố gắng,"Em cần tích cực hơn trong các hoạt động tập thể."
Yêu nước,Cần cố gắng,"Em cần có ý thức hơn trong việc bảo vệ môi trường và của công."
Yêu nước,Cần cố gắng,"Em còn thụ động trong các phong trào thi đua của lớp."
Nhân ái,Tốt,"Em hòa đồng, biết giúp đỡ và chia sẻ với bạn bè."
Nhân ái,Tốt,"Em vui vẻ hòa đồng, tôn trọng người lớn tuổi, yêu quý bạn bè."
Nhân ái,Tốt,"Em có tấm lòng nhân ái, sẵn sàng giúp đỡ người gặp hoàn cảnh khó khăn."
Nhân ái,Tốt,"Em luôn biết cảm thông, chia sẻ và sẵn lòng giúp đỡ mọi người khi gặp khó khăn."
Nhân ái,Tốt,"Em biết nhường nhịn, yêu thương em nhỏ và giúp đỡ người già."
Nhân ái,Đạt,"Em cư xử hòa nhã với bạn bè và thầy cô."
Nhân ái,Đạt,"Em biết yêu thương và giúp đỡ bạn bè phù hợp với hoàn cảnh."
Nhân ái,Đạt,"Em biết chia sẻ buồn vui cùng bạn bè trong lớp."
Nhân ái,Đạt,"Em biết vâng lời cha mẹ, thầy cô giáo và lễ phép với mọi người."
Nhân ái,Cần cố gắng,"Em cần chủ động hơn trong việc giúp đỡ bạn bè."
Nhân ái,Cần cố gắng,"Em còn hay trêu chọc bạn, cần đoàn kết hơn."
Nhân ái,Cần cố gắng,"Em chưa biết kiềm chế cảm xúc, cần hòa nhã hơn với bạn."
Chăm chỉ,Tốt,"Em đi học đầy đủ, đúng giờ, tự giác thực hiện nhiệm vụ học tập."
Chăm chỉ,Tốt,"Em chăm học và hoàn thành xuất sắc nhiệm vụ học tập."
Chăm chỉ,Tốt,"Em có ý thức học hỏi, tự giác làm bài và có ý chí vượt khó."
Chăm chỉ,Tốt,"Em luôn nỗ lực vượt qua khó khăn để hoàn thành tốt các nhiệm vụ học tập."
Chăm chỉ,Tốt,"Em thường xuyên tìm tòi tài liệu mới để mở rộng kiến thức môn học."
Chăm chỉ,Đạt,"Em có ý thức học tập và hoàn thành nhiệm vụ được giao."
Chăm chỉ,Đạt,"Em chăm chú nghe giảng và làm bài tập đầy đủ."
Chăm chỉ,Đạt,"Em hoàn thành các bài học trên lớp."
Chăm chỉ,Đạt,"Em có tinh thần vươn lên trong học tập, hoàn thành tốt các yêu cầu của giáo viên."
Chăm chỉ,Cần cố gắng,"Em cần tập trung hơn trong học tập."
Chăm chỉ,Cần cố gắng,"Em chưa chăm học đều, thường xuyên quên sách vở."
Chăm chỉ,Cần cố gắng,"Em còn hay làm việc riêng trong giờ học, cần chú ý hơn."
Trung thực,Tốt,"Em trung thực trong học tập và sinh hoạt, không tham của rơi."
Trung thực,Tốt,"Em luôn thật thà, biết nhận lỗi và sửa lỗi khi phạm lỗi."
Trung thực,Tốt,"Em không nói dối, trung thực trong kiểm tra và đánh giá."
Trung thực,Tốt,"Em có tinh thần trung thực, luôn báo cáo đúng kết quả học tập và sinh hoạt."
Trung thực,Tốt,"Em biết đấu tranh with những hành vi thiếu trung thực xung quanh mình."
Trung thực,Đạt,"Em thực hiện tương đối tốt các yêu cầu học tập, biết nhận lỗi."
Trung thực,Đạt,"Em có ý thức nói thật, không lấy đồ của bạn."
Trung thực,Đạt,"Em biết đồng tình with những hành vi trung thực trong cuộc sống."
Trung thực,Cần cố gắng,"Em cần mạnh dạn and trung thực hơn trong học tập."
Trung thực,Cần cố gắng,"Em chưa tự giác nhận lỗi khi làm sai."
Trung thực,Cần cố gắng,"Em cần được nhắc nhở thường xuyên về việc trung thực trong kiểm tra."
Trách nhiệm,Tốt,"Em có trách nhiệm với công việc được giao ở trường, lớp."
Trách nhiệm,Tốt,"Em có ý thức bảo vệ của công, giữ gìn vệ sinh môi trường sạch sẽ."
Trách nhiệm,Tốt,"Em luôn hoàn thành xuất sắc các nhiệm vụ do nhóm phân công."
Trách nhiệm,Tốt,"Em luôn có tinh thần trách nhiệm cao, hoàn thành đúng hạn các nhiệm vụ được giao."
Trách nhiệm,Tốt,"Em biết sắp xếp nhà cửa, góc học tập gọn gàng, ngăn nắp và khoa học."
Trách nhiệm,Đạt,"Em thực hiện tương đối đầy đủ nhiệm vụ được giao."
Trách nhiệm,Đạt,"Em có tham gia trực nhật và giữ vệ sinh lớp."
Trách nhiệm,Đạt,"Em biết bảo quản đồ dùng học tập cá nhân và của tập thể."
Trách nhiệm,Cần cố gắng,"Em cần nâng cao tinh thần trách nhiệm với tập thể."
Trách nhiệm,Cần cố gắng,"Em hay quên trực nhật, cần có ý thức hơn đối với công việc chung."
Tự chủ và tự học,Tốt,"Em có ý thức tự học và tự chủ trong mọi vấn đề."
Tự chủ và tự học,Tốt,"Em có khả năng tự thực hiện tốt các nhiệm vụ học tập."
Tự chủ và tự học,Tốt,"Em tự giác làm bài không cần nhắc nhở, biết tự điều chỉnh hành vi."
Tự chủ and tự học,Tốt,"Em biết tự lập lộ trình học tập and hoàn thành kế hoạch cá nhân một cách sáng tạo."
Tự chủ and tự học,Tốt,"Em chủ động tìm kiếm sự hỗ trợ khi cần thiết để giải quyết vấn đề cá nhân."
Tự chủ and tự học,Đạt,"Em bước đầu biết tự thực hiện nhiệm vụ học tập."
Tự chủ and tự học,Đạt,"Em biết tự phục vụ nhu cầu bản thân ở mức cơ bản."
Tự chủ and tự học,Đạt,"Em hoàn thành các nhiệm vụ học tập with sự giúp đỡ from thầy cô and bạn bè."
Tự chủ and tự học,Cần cố gắng,"Em chưa có ý thức tự học, cần nhắc nhở nhiều."
Tự chủ and tự học,Cần cố gắng,"Em còn phụ thuộc vào người khác, chưa tự giác hoàn thành bài làm."
Giao tiếp and hợp tác,Tốt,"Em trình bày rõ ràng, ngắn gọn nội dung cần trao đổi."
Giao tiếp and hợp tác,Tốt,"Em biết trao đổi ý kiến cùng bạn rất tốt, mạnh dạn đưa ra ý kiến cá nhân."
Giao tiếp and hợp tác,Tốt,"Em có khả năng phối hợp with bạn bè khi làm việc nhóm rất hiệu quả."
Giao tiếp and hợp tác,Tốt,"Em luôn chủ động lắng nghe and phối hợp ăn ý with các thành viên trong nhóm."
Giao tiếp and hợp tác,Tốt,"Em có phong thái tự tin khi giao tiếp and biết cách thuyết phục người nghe."
Giao tiếp and hợp tác,Đạt,"Em biết cách giao tiếp cơ bản and hợp tác with nhóm."
Giao tiếp and hợp tác,Đạt,"Em biết lắng nghe and chia sẻ ý kiến with các bạn trong hoạt động tập thể."
Giao tiếp and hợp tác,Đạt,"Em hòa nhã, cởi mở khi trò chuyện cùng thầy cô and bạn bè."
Giao tiếp and hợp tác,Cần cố gắng,"Em chưa mạnh dạn trong giao tiếp and hợp tác."
Giao tiếp and hợp tác,Cần cố gắng,"Em còn thụ động, ít phát biểu ý kiến cá nhân trong quá trình làm việc nhóm."
Giải quyết vấn đề,Tốt,"Em biết thu thập thông tin and tình huống, nhận ra những vấn đề đơn giản."
Giải quyết vấn đề,Tốt,"Em có năng lực giải quyết vấn đề, sáng tạo trong học tập."
Giải quyết vấn đề,Tốt,"Em xử lý tình huống nhanh, chính xác and có ý tưởng độc đáo."
Giải quyết vấn đề,Tốt,"Em luôn linh hoạt trong tư duy and tìm ra nhiều giải pháp hiệu quả cho bài tập khó."
Giải quyết vấn đề,Tốt,"Em có khả năng tự nhận xét and đánh giá mức độ hoàn thành nhiệm vụ của bản thân."
Giải quyết vấn đề,Đạt,"Em có khả năng giải quyết các vấn đề học tập cơ bản."
Giải quyết vấn đề,Đạt,"Em bước đầu biết cách thức giải quyết vấn đề theo hướng dẫn của giáo viên."
Giải quyết vấn đề,Đạt,"Em biết xử lý các tình huống thực tế thường gặp một cách phù hợp."
Giải quyết vấn đề,Cần cố gắng,"Em xử lí tình huống còn lúng túng, chưa linh hoạt."
Giải quyết vấn đề,Cần cố gắng,"Em cần cố gắng suy nghĩ độc lập hơn trước khi nhờ sự hỗ trợ của giáo viên."
Ngôn ngữ,Tốt,"Đọc to, rõ ràng, lưu loát. Câu văn ngắn gọn, dễ hiểu."
Ngôn ngữ,Tốt,"Viết chính tả chính xác. Chữ viết đều, đúng nét, trình bày sạch đẹp."
Ngôn ngữ,Tốt,"Em có vốn từ phong phú, diễn đạt câu trọn vẹn, sinh động, có hình ảnh hay."
Ngôn ngữ,Tốt,"Kỹ năng đọc hiểu rất tốt, biết diễn đạt ý tưởng qua bài viết một cách sâu sắc and mạch lạc."
Ngôn ngữ,Tốt,"Em biết sử dụng linh hoạt các kiểu câu and từ ngữ để viết đoạn văn cảm xúc."
Ngôn ngữ,Đạt,"Đọc bài khá lưu loát, chữ viết tương đối rõ ràng, sạch sẽ."
Ngôn ngữ,Đạt,"Hiểu nội dung bài đọc and trả lời tốt các câu hỏi của bài học."
Ngôn ngữ,Đạt,"Em viết được câu có đủ thành phần and diễn đạt được ý của mình."
Ngôn ngữ,Đạt,"Em có tốc độ đọc đạt yêu cầu, viết đúng chính tả with tốc độ phù hợp."
Ngôn ngữ,Cần cố gắng,"Đọc còn đánh vần chậm, mắc nhiều lỗi chính tả."
Ngôn ngữ,Cần cố gắng,"Kĩ năng đọc hiểu văn bản còn hạn chế, diễn đạt câu chưa rõ ý."
Ngôn ngữ,Cần cố gắng,"Chữ viết chưa đúng mẫu, cần rèn luyện thêm về nét chữ and trình bày."
Tính toán,Tốt,"Nắm vững kiến thức đã học. Tính toán thành thạo, giải toán đúng and nhanh."
Tính toán,Tốt,"Thông minh, có trí nhớ tốt, tính toán nhanh nhẹn and chính xác."
Tính toán,Tốt,"Có tư duy toán tốt, giải toán có lời văn rất sáng tạo and trình bày khoa học."
Tính toán,Tốt,"Phát hiện nhanh các quy luật toán học and áp dụng hiệu quả vào thực hành."
Tính toán,Tốt,"Em nắm chắc kiến thức về hình học and biết vận dụng linh hoạt vào thực tiễn."
Tính toán,Đạt,"Hiểu bài, nắm được các kiến thức toán học cơ bản."
Tính toán,Đạt,"Thực hiện được các phép tính nhưng giải toán có lời văn cần rèn luyện thêm."
Tính toán,Đạt,"Em biết cách xem ngày giờ, tháng and thực hiện tốt các bài tập đo lường."
Tính toán,Đạt,"Tính toán tương đối tốt, hoàn thành đầy đủ các bài tập được giao."
Tính toán,Cần cố gắng,"Chưa nắm chắc kiến thức cơ bản, kĩ năng tính toán còn nhiều hạn chế."
Tính toán,Cần cố gắng,"Tiếp thu bài chậm, tính toán còn hay nhầm lẫn, cần cẩn thận hơn."
Khoa học,Tốt,"Hiểu bài, vận dụng tốt các kiến thức đã học vào thực tế cuộc sống."
Khoa học,Tốt,"Chủ động nắm bắt and ghi nhớ kiến thức tốt, ham học hỏi tìm tòi khám phá."
Khoa học,Tốt,"Có vốn hiểu biết phong phú về tự nhiên, xã hội and bảo vệ môi trường."
Khoa học,Tốt,"Sáng tạo trong việc thực hiện các dự án nhỏ về khoa học and tìm hiểu thế giới."
Khoa học,Đạt,"Hoàn thành tốt nội dung and yêu cầu kiến thức môn học."
Khoa học,Đạt,"Nắm được kiến thức cơ bản and biết vận dụng bài học vào bảo vệ sức khỏe bản thân."
Khoa học,Đạt,"Biết quan sát tranh ảnh để rút ra kiến thức bài học một cách chính xác."
Khoa học,Cần cố gắng,"Cần tích cực chủ động hơn trong việc tìm hiểu các hiện tượng tự nhiên."
Khoa học,Cần cố gắng,"Vốn hiểu biết về khoa học còn hạn chế, cần rèn luyện thêm kỹ năng quan sát."
Thẩm mĩ,Tốt,"Có khiếu thẩm mĩ, thực hành khéo tay, sản phẩm sáng tạo and đẹp mắt."
Thẩm mĩ,Tốt,"Hát hay, biểu diễn tự nhiên, tự tin thể hiện cảm xúc qua bài hát."
Thẩm mĩ,Tốt,"Em vẽ tranh có bố cục hài hòa, màu sắc sinh động and giàu ý tưởng."
Thẩm mĩ,Tốt,"Hát đúng nhịp, giai điệu and biết kết hợp vận động phụ họa nhịp nhàng."
Thẩm mĩ,Đạt,"Hoàn thành tốt các sản phẩm mĩ thuật theo yêu cầu."
Thẩm mĩ,Đạt,"Thuộc lời bài hát, hát đúng lời ca and rõ tiếng."
Thẩm mĩ,Đạt,"Biết sử dụng màu sắc phù hợp để trang trí sản phẩm của mình."
Thẩm mĩ,Cần cố gắng,"Cần nhanh nhẹn and cẩn thận hơn trong các thao tác mĩ thuật."
Thẩm mĩ,Cần cố gắng,"Hát còn nhỏ, chưa tự tin khi tham gia biểu diễn trước lớp."
Thể chất,Tốt,"Hoàn thành tốt lượng vận động, tham gia tích cực and sôi nổi các trò chơi."
Thể chất,Tốt,"Thực hiện chuẩn xác các động tác thể dục, thao tác nhanh nhẹn, dứt khoát."
Thể chất,Tốt,"Em có ý thức rèn luyện thân thể tốt, phát huy được năng khiếu thể thao."
Thể chất,Tốt,"Nhanh nhẹn, có tinh thần cầu tiến and tích cực giúp đỡ bạn cùng tập luyện."
Thể chất,Đạt,"Em thực hiện được các động tác thể dục cơ bản theo yêu cầu."
Thể chất,Đạt,"Có ý thức rèn luyện and tham gia các trò chơi vận động cùng nhóm."
Thể chất,Đạt,"Hoàn thành các bài tập tư thế and kỹ năng vận động cơ bản."
Thể chất,Cần cố gắng,"Chưa tích cực tham gia vận động, cần tập trung hơn trong giờ học."
Thể chất,Cần cố gắng,"Cần rèn luyện thêm về kỹ thuật động tác để đạt kết quả tốt hơn."
Công nghệ,Tốt,"Em có hiểu biết tốt về thế giới công nghệ and biết vận dụng thực hành hiệu quả."
Công nghệ,Tốt,"Lắp ráp mô hình kĩ thuật sáng tạo, chắc chắn and trình bày đẹp."
Công nghệ,Tốt,"Sử dụng các dụng cụ an toàn, thành thạo and bảo quản thiết bị rất tốt."
Công nghệ,Đạt,"Em có kiến thức nền tảng về môn học and hoàn thành các bài thực hành."
Công nghệ,Đạt,"Sản phẩm hoàn thành đạt tiêu chuẩn chất lượng theo yêu cầu bài học."
Công nghệ,Cần cố gắng,"Thao tác lắp ráp còn lúng túng, cần rèn luyện thêm kỹ năng thực hành."
Tin học,Tốt,"Em thao tác tốt with chuột and bàn phím, thực hiện nhanh and chính xác."
Tin học,Tốt,"Sử dụng linh hoạt các phần mềm học tập and có ý thức an toàn mạng."
Tin học,Tốt,"Em biết phân biệt giữa nháy chuột and nháy đúp chuột khi thao tác trên máy."
Tin học,Tốt,"Em gõ bàn phím 10 ngón đúng quy tắc and nhập văn bản nhanh, chính xác."
Tin học,Tốt,"Có khả năng khai thác phần mềm để ứng dụng vào các môn học khác hiệu quả."
Tin học,Đạt,"Em biết cách khởi động and thoát khỏi phần mềm đúng quy trình."
Tin học,Đạt,"Hoàn thành tốt các bài thi and thực hành trên máy ở mức yêu cầu."
Tin học,Đạt,"Biết cách sử dụng máy tính an toàn and giữ gìn thiết bị của lớp."
Tin học,Cần cố gắng,"Em cần thực hành nâng cao kỹ năng sử dụng chuột and bàn phím."
Tin học,Cần cố gắng,"Cần tập trung hơn để nắm vững các lệnh điều khiển trong phần mềm."`;

const criteriaMapping: Record<string, string[]> = {
  'Yêu nước': ['yêu nước'],
  'Nhân ái': ['nhân ái'],
  'Chăm chỉ': ['chăm chỉ'],
  'Trung thực': ['trung thực'],
  'Trách nhiệm': ['trách nhiệm'],
  'Tự chủ và tự học': ['tự chủ', 'tự học'],
  'Giao tiếp và hợp tác': ['giao tiếp', 'hợp tác'],
  'Giải quyết vấn đề': ['giải quyết vấn đề', 'gqvđ', 'sáng tạo'],
  'Ngôn ngữ': ['ngôn ngữ'],
  'Tính toán': ['tính toán'],
  'Khoa học': ['khoa học'],
  'Thẩm mĩ': ['thẩm mĩ', 'thẩm mỹ'],
  'Thể chất': ['thể chất'],
  'Công nghệ': ['công nghệ'],
  'Tin học': ['tin học']
};

const categoryMap: Record<string, 'chung' | 'dacThu' | 'phamChat'> = {
  'Tự chủ và tự học': 'chung',
  'Giao tiếp và hợp tác': 'chung',
  'Giải quyết vấn đề': 'chung',
  'Ngôn ngữ': 'dacThu',
  'Tính toán': 'dacThu',
  'Khoa học': 'dacThu',
  'Thẩm mĩ': 'dacThu',
  'Thể chất': 'dacThu',
  'Công nghệ': 'dacThu',
  'Tin học': 'dacThu',
  'Yêu nước': 'phamChat',
  'Nhân ái': 'phamChat',
  'Chăm chỉ': 'phamChat',
  'Trung thực': 'phamChat',
  'Trách nhiệm': 'phamChat'
};

const levelMapping: Record<string, string> = {
  'T': 'Tốt',
  'Đ': 'Đạt',
  'C': 'Cần cố gắng',
  'TỐT': 'Tốt',
  'ĐẠT': 'Đạt',
  'CẦN CỐ GẮNG': 'Cần cố gắng',
  'HTT': 'Tốt',
  'H': 'Đạt',
  'CHT': 'Cần cố gắng'
};

interface Student {
  id: string;
  studentId: string;
  name: string;
  originalGrades: Record<string, string>;
  individualComments: Record<string, string>;
  comments: {
    chung: string;
    dacThu: string;
    phamChat: string;
  };
}

interface ParsedBank {
  [criteria: string]: {
    [level: string]: string[];
  };
}

// --- Helper Functions ---

function parseCSVRow(text: string): string[] {
  let ret = [''], i = 0, s = true;
  for (let l = text.length; i < l; i++) {
    let c = text[i];
    if (c === '"') {
      s = !s;
    } else if (c === ',' && s) {
      ret.push('');
    } else {
      ret[ret.length - 1] += c;
    }
  }
  return ret.map(val => val.replace(/^"|"$/g, '').trim());
}

// --- Main Component ---

export default function App() {
  const [activeTab, setActiveTab] = useState<'generator' | 'settings'>('generator');
  const [commentBankText, setCommentBankText] = useState(() => {
    if (typeof window !== 'undefined') {
      const saved = localStorage.getItem('auto_comment_bank');
      return saved || defaultCommentBank;
    }
    return defaultCommentBank;
  });
  const [parsedBank, setParsedBank] = useState<ParsedBank>({});
  const [students, setStudents] = useState<Student[]>([]);
  const [fileName, setFileName] = useState('');
  const [statusMsg, setStatusMsg] = useState({ text: '', type: '' });
  const [gradeBlock, setGradeBlock] = useState<'k12' | 'k345'>(() => {
    if (typeof window !== 'undefined') {
      return (localStorage.getItem('auto_grade_block') as any) || 'k12';
    }
    return 'k12';
  }); 
  const [semester, setSemester] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      return localStorage.getItem('auto_semester') || 'Học kỳ 1';
    }
    return 'Học kỳ 1';
  });
  const [customSemester, setCustomSemester] = useState('');
  const [isCustomSemester, setIsCustomSemester] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [editingStudent, setEditingStudent] = useState<Student | null>(null);

  const saveCommentBank = () => {
    localStorage.setItem('auto_comment_bank', commentBankText);
    setStatusMsg({ text: 'Đã lưu thay đổi ngân hàng câu nhận xét vào bộ nhớ trình duyệt.', type: 'success' });
  };

  const resetCommentBank = () => {
    if (confirm('Bạn có chắc chắn muốn phục hồi ngân hàng câu về mặc định? Mọi thay đổi hiện tại sẽ bị xóa.')) {
      setCommentBankText(defaultCommentBank);
      localStorage.removeItem('auto_comment_bank');
      setStatusMsg({ text: 'Đã phục hồi ngân hàng câu về mặc định.', type: 'info' });
    }
  };

  const downloadCommentBank = () => {
    const blob = new Blob([commentBankText], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'ngan_hang_nhan_xet_custom.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const importCommentBank = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv,.txt';
    input.onchange = (e: any) => {
      const file = e.target.files[0];
      if (!file) return;
      
      const ext = file.name.split('.').pop()?.toLowerCase();
      if (!['csv', 'txt'].includes(ext || '')) {
        setStatusMsg({ text: "Vui lòng chọn file .csv hoặc .txt định dạng UTF-8.", type: 'error' });
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        const text = evt.target?.result as string;
        if (text) {
          // Basic validation: check if it looks like CSV
          if (!text.includes(',') && !text.includes('\n')) {
            setStatusMsg({ text: "Nội dung file không đúng định dạng ngân hàng câu (CSV).", type: 'error' });
            return;
          }
          setCommentBankText(text);
          setStatusMsg({ text: 'Đã nạp ngân hàng câu từ file thành công. Nhấn "Lưu Thay Đổi" để duy trì lâu dài.', type: 'info' });
        }
      };
      reader.onerror = () => setStatusMsg({ text: "Lỗi khi đọc file ngân hàng câu.", type: 'error' });
      reader.readAsText(file);
    };
    input.click();
  };

  useEffect(() => {
    localStorage.setItem('auto_grade_block', gradeBlock);
  }, [gradeBlock]);

  useEffect(() => {
    localStorage.setItem('auto_semester', semester);
    // Sync isCustomSemester if semester doesn't match defaults
    if (semester !== 'Học kỳ 1' && semester !== 'Học kỳ 2' && semester !== '') {
      setIsCustomSemester(true);
      setCustomSemester(semester);
    } else if (semester === 'Học kỳ 1' || semester === 'Học kỳ 2') {
      setIsCustomSemester(false);
      setCustomSemester('');
    }
  }, [semester]);

  useEffect(() => {
    const lines = commentBankText.split('\n');
    const bank: ParsedBank = {};
    Object.keys(criteriaMapping).forEach(c => { bank[c] = {}; });

    lines.slice(1).forEach(line => {
      if (!line.trim()) return;
      const [criteria, level, ...commentParts] = parseCSVRow(line);
      let comment = commentParts.join(',').replace(/^"|"$/g, '').trim();
      
      if (comment && comment.length > 0) {
        if (comment === comment.toUpperCase()) {
          comment = comment.charAt(0).toUpperCase() + comment.slice(1).toLowerCase();
        } else {
          comment = comment.charAt(0).toUpperCase() + comment.slice(1);
        }
      }
      
      if (criteria && level && bank[criteria]) {
        if (!bank[criteria][level]) {
          bank[criteria][level] = [];
        }
        bank[criteria][level].push(comment);
      }
    });
    setParsedBank(bank);
  }, [commentBankText]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    // Check file extension
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(ext || '')) {
      setStatusMsg({ 
        text: "Định dạng file không hỗ trợ. Vui lòng sử dụng file Excel (.xlsx, .xls) hoặc CSV.", 
        type: 'error' 
      });
      return;
    }

    setFileName(file.name);
    setStudents([]);
    setStatusMsg({ text: 'Đang đọc và phân tích dữ liệu file...', type: 'info' });
    
    const reader = new FileReader();
    reader.onerror = () => {
      setStatusMsg({ text: "Lỗi hệ thống khi đọc file. Vui lòng thử lại hoặc sử dụng trình duyệt khác.", type: 'error' });
    };
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          setStatusMsg({ text: "File Excel không có trang tính (Sheet) nào.", type: 'error' });
          return;
        }
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "", raw: false });
        processDataRows(rows);
      } catch (error) {
        console.error("XLSX Read Error:", error);
        setStatusMsg({ text: "Lỗi cấu trúc file Excel. Vui lòng đảm bảo file không bị lỗi hoặc có mật khẩu bảo vệ.", type: 'error' });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const processDataRows = (rows: any[][]) => {
    if (rows.length === 0) {
      setStatusMsg({ 
        text: "File rỗng hoặc không có dữ liệu. Vui lòng kiểm tra lại nội dung file Excel.", 
        type: 'error' 
      });
      return;
    }

    let colIndices = { name: -1, studentId: -1 };
    let foundCriteriaCols: Record<string, number> = {};
    let dataStartIndex = -1;

    // Search headers in the first 25 rows
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      const cols = rows[i].map(c => String(c || '').toLowerCase().trim());
      
      cols.forEach((colStr, index) => {
        if (!colStr) return;
        
        // Match Name column
        if (colStr.includes('họ và tên') || colStr.includes('họ tên') || colStr === 'tên học sinh' || colStr === 'họ và tên học sinh') {
          if (colIndices.name === -1) colIndices.name = index;
        }
        
        // Match Student ID column
        if (colStr.includes('mã học sinh') || colStr.includes('mã hs') || colStr.includes('mã định danh') || colStr.includes('mã số học sinh')) {
          if (colIndices.studentId === -1) colIndices.studentId = index;
        }

        // Match Criteria columns
        Object.keys(criteriaMapping).forEach(mainCriteria => {
          const keywords = criteriaMapping[mainCriteria];
          if (keywords.some(kw => colStr.includes(kw))) {
            // Prefer columns that are shorter or exact matches if multiple exist
            if (foundCriteriaCols[mainCriteria] === undefined) {
              foundCriteriaCols[mainCriteria] = index;
            }
          }
        });
      });

      // Try to detect where data starts (first row with a number in first column)
      if (cols.length > 0 && String(cols[0]).trim() !== '' && /^\d+$/.test(String(cols[0]).trim())) {
        if (dataStartIndex === -1) dataStartIndex = i;
      }
    }

    // specific errors for missing columns
    if (colIndices.name === -1) {
      setStatusMsg({ 
        text: "Không tìm thấy cột 'Họ và tên'. Vui lòng đảm bảo file của bạn có cột tiêu đề là 'Họ và tên' hoặc 'Họ tên'.", 
        type: 'error' 
      });
      return;
    }

    const foundCount = Object.keys(foundCriteriaCols).length;
    if (foundCount === 0) {
      setStatusMsg({ 
        text: "Không nhận diện được các cột Năng lực/Phẩm chất (như Ngôn ngữ, Tính toán, Yêu nước...). Vui lòng kiểm tra tiêu đề các cột trong file.", 
        type: 'error' 
      });
      return;
    }

    if (foundCount < 5) {
      setStatusMsg({ 
        text: `Cảnh báo: Chỉ tìm thấy ${foundCount} cột năng lực. Kết quả có thể không đầy đủ. Tiếp tục xử lý...`, 
        type: 'info' 
      });
    }

    const results: Student[] = [];
    const startIdx = dataStartIndex !== -1 ? dataStartIndex : 1; // Fallback to 1 if not detected

    try {
      for (let i = startIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length <= colIndices.name || !row[colIndices.name]) continue;

        const name = String(row[colIndices.name]).trim();
        // Skip obvious header/footer text that might be detected as names
        if (!name || name.length < 2 || name.toLowerCase().includes('tổng số') || name.toLowerCase().includes('trung bình') || name.toLowerCase().includes('họ và tên')) continue;
        
        const studentId = colIndices.studentId !== -1 ? String(row[colIndices.studentId] || '').trim() : '';

        const originalGrades: Record<string, string> = {};
        const individualComments: Record<string, string> = {}; 
        const studentIdxInSequence = results.length; 
        
        Object.keys(foundCriteriaCols).forEach(criteria => {
          const colIdx = foundCriteriaCols[criteria];
          let rawGrade = row[colIdx] ? String(row[colIdx]).trim() : '';
          let gradeStr = rawGrade.toUpperCase();
          originalGrades[criteria] = gradeStr;
          const level = levelMapping[gradeStr] || gradeStr;
          
          const possibleComments = parsedBank[criteria] && parsedBank[criteria][level];
          if (possibleComments && possibleComments.length > 0) {
            individualComments[criteria] = possibleComments[studentIdxInSequence % possibleComments.length];
          } else {
            let fallback = rawGrade;
            if (fallback && fallback.length > 0) {
              if (fallback === fallback.toUpperCase()) {
                fallback = fallback.charAt(0).toUpperCase() + fallback.slice(1).toLowerCase();
              } else {
                fallback = fallback.charAt(0).toUpperCase() + fallback.slice(1);
              }
            }
            individualComments[criteria] = fallback || "Chưa có nhận xét cho mức này.";
          }
        });

        const grouped = generateGroupedCommentsInternal(individualComments);

        results.push({
          id: `std_${i}_${Date.now()}`,
          studentId,
          name,
          originalGrades,
          individualComments,
          comments: grouped
        });
      }

      if (results.length === 0) {
        setStatusMsg({ 
          text: "Không tìm thấy dữ liệu học sinh hợp lệ sau dòng tiêu đề. Vui lòng kiểm tra cấu trúc file.", 
          type: 'error' 
        });
        return;
      }

      setStudents(results);
      setStatusMsg({ 
        text: `Xử lý thành công! Đã nạp ${results.length} học sinh.`, 
        type: 'success' 
      });
    } catch (err) {
      console.error("Processing error:", err);
      setStatusMsg({ 
        text: "Có lỗi xảy ra trong quá trình phân tích dữ liệu từng dòng. Vui lòng kiểm tra định dạng file.", 
        type: 'error' 
      });
    }
  };

  const generateGroupedCommentsInternal = (individualComments: Record<string, string>) => {
    let grouped = { chung: [] as string[], dacThu: [] as string[], phamChat: [] as string[] };
    
    Object.keys(individualComments).forEach(criteria => {
      const comment = individualComments[criteria];
      const cat = categoryMap[criteria];
      
      // Filter out tech for k12 if needed (as per user's logic)
      if (gradeBlock === 'k12' && (criteria === 'Công nghệ' || criteria === 'Tin học')) return;

      if (cat && comment && comment.length > 2) {
        grouped[cat].push(comment);
      }
    });

    const formatString = (arr: string[]) => {
      if (arr.length === 0) return "";
      let str = arr.join(' ');
      return str.charAt(0).toUpperCase() + str.slice(1);
    };

    return {
      chung: formatString(grouped.chung),
      dacThu: formatString(grouped.dacThu),
      phamChat: formatString(grouped.phamChat)
    };
  };

  const openEditModal = (student: Student) => {
    setEditingStudent(JSON.parse(JSON.stringify(student)));
  };

  const closeEditModal = () => {
    setEditingStudent(null);
  };

  const handleEditIndividualComment = (criteria: string, value: string) => {
    if (!editingStudent) return;
    setEditingStudent(prev => {
      if (!prev) return null;
      const updatedIndividual = { ...prev.individualComments, [criteria]: value };
      return {
        ...prev,
        individualComments: updatedIndividual
      };
    });
  };

  const handleEditGroupedComment = (cat: 'chung' | 'dacThu' | 'phamChat', value: string) => {
    if (!editingStudent) return;
    setEditingStudent(prev => {
      if (!prev) return null;
      return {
        ...prev,
        comments: { ...prev.comments, [cat]: value }
      };
    });
  };

  const regenerateGroupedForEditingStudent = () => {
    if (!editingStudent) return;
    const newGrouped = generateGroupedCommentsInternal(editingStudent.individualComments);
    setEditingStudent(prev => {
      if (!prev) return null;
      return {
        ...prev,
        comments: newGrouped
      };
    });
  };

  const saveEditedStudent = () => {
    if (!editingStudent) return;
    setStudents(prevStudents => 
      prevStudents.map(std => std.id === editingStudent.id ? editingStudent : std)
    );
    closeEditModal();
  };

  const exportToStandardTemplateExcel = () => {
    if (students.length === 0) return;
    
    let wsData: any[][] = [];
    let merges: XLSX.Range[] = [];
    let cols: XLSX.ColInfo[] = [];

    if (gradeBlock === 'k12') {
      wsData = [
        [
          "STT", "Mã học sinh", "Họ và tên ", "Học kì", 
          "Năng lực chung", "", "", "", 
          "Năng lực đặc thù", "", "", "", "", "", 
          "Phẩm chất", "", "", "", "", ""
        ],
        [
          "", "", "", "", 
          "Nhận xét về năng lực", "Tự chủ và tự học", "Giao tiếp và hợp tác", "GQVĐ và sáng tạo", 
          "Nhận xét về năng lực đặc thù", "Ngôn ngữ", "Tính toán", "Khoa học", "Thẩm mĩ", "Thể chất", 
          "Nhận xét về phẩm chất", "Yêu nước", "Nhân ái", "Chăm chỉ", "Trung thực", "Trách nhiệm"
        ]
      ];

      students.forEach((std, idx) => {
        wsData.push([
          idx + 1,
          std.studentId || "",
          std.name,
          semester, 
          std.comments.chung,
          std.individualComments["Tự chủ và tự học"] || "",
          std.individualComments["Giao tiếp và hợp tác"] || "",
          std.individualComments["Giải quyết vấn đề"] || "",
          std.comments.dacThu,
          std.individualComments["Ngôn ngữ"] || "",
          std.individualComments["Tính toán"] || "",
          std.individualComments["Khoa học"] || "",
          std.individualComments["Thẩm mĩ"] || "",
          std.individualComments["Thể chất"] || "",
          std.comments.phamChat,
          std.individualComments["Yêu nước"] || "",
          std.individualComments["Nhân ái"] || "",
          std.individualComments["Chăm chỉ"] || "",
          std.individualComments["Trung thực"] || "",
          std.individualComments["Trách nhiệm"] || ""
        ]);
      });

      merges = [
        { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },  { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, 
        { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },  { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, 
        { s: { r: 0, c: 4 }, e: { r: 0, c: 7 } }, 
        { s: { r: 0, c: 8 }, e: { r: 0, c: 13 } },
        { s: { r: 0, c: 14 }, e: { r: 0, c: 19 } }
      ];

      cols = [
        { wch: 5 }, { wch: 15 }, { wch: 22 }, { wch: 10 }, 
        { wch: 50 }, { wch: 30 }, { wch: 30 }, { wch: 30 },
        { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, 
        { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }
      ];

    } else {
      wsData = [
        [
          "STT", "Mã học sinh", "Họ và tên ", "Học kì", 
          "Năng lực chung", "", "", "", 
          "Năng lực đặc thù", "", "", "", "", "", "", "", 
          "Phẩm chất", "", "", "", "", ""
        ],
        [
          "", "", "", "", 
          "Nhận xét về năng lực", "Tự chủ và tự học", "Giao tiếp và hợp tác", "GQVĐ và sáng tạo", 
          "Nhận xét về năng lực đặc thù", "Ngôn ngữ", "Tính toán", "Khoa học", "Thẩm mĩ", "Thể chất", "Công nghệ", "Tin học",
          "Nhận xét về phẩm chất", "Yêu nước", "Nhân ái", "Chăm chỉ", "Trung thực", "Trách nhiệm"
        ]
      ];

      students.forEach((std, idx) => {
        wsData.push([
          idx + 1,
          std.studentId || "",
          std.name,
          semester, 
          std.comments.chung,
          std.individualComments["Tự chủ và tự học"] || "",
          std.individualComments["Giao tiếp và hợp tác"] || "",
          std.individualComments["Giải quyết vấn đề"] || "",
          std.comments.dacThu,
          std.individualComments["Ngôn ngữ"] || "",
          std.individualComments["Tính toán"] || "",
          std.individualComments["Khoa học"] || "",
          std.individualComments["Thẩm mĩ"] || "",
          std.individualComments["Thể chất"] || "",
          std.individualComments["Công nghệ"] || "",
          std.individualComments["Tin học"] || "",
          std.comments.phamChat,
          std.individualComments["Yêu nước"] || "",
          std.individualComments["Nhân ái"] || "",
          std.individualComments["Chăm chỉ"] || "",
          std.individualComments["Trung thực"] || "",
          std.individualComments["Trách nhiệm"] || ""
        ]);
      });

      merges = [
        { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },  { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, 
        { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },  { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, 
        { s: { r: 0, c: 4 }, e: { r: 0, c: 7 } }, 
        { s: { r: 0, c: 8 }, e: { r: 0, c: 15 } }, 
        { s: { r: 0, c: 16 }, e: { r: 0, c: 21 } } 
      ];

      cols = [
        { wch: 5 }, { wch: 15 }, { wch: 22 }, { wch: 10 }, 
        { wch: 50 }, { wch: 30 }, { wch: 30 }, { wch: 30 },
        { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, 
        { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }
      ];
    }

    const worksheet = XLSX.utils.aoa_to_sheet(wsData);
    worksheet['!merges'] = merges;
    worksheet['!cols'] = cols;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Tong_Hop");
    
    const outFileName = gradeBlock === 'k12' ? "Ket_Qua_Nhan_Xet_Khoi_1_2.xlsx" : "Ket_Qua_Nhan_Xet_Khoi_3_4_5.xlsx";
    XLSX.writeFile(workbook, outFileName);
  };

  return (
    <div className="min-h-screen bg-neutral-50 font-sans text-neutral-800" style={{ fontFamily: "'Inter', sans-serif" }}>
      <div className="max-w-[1600px] mx-auto min-h-screen flex flex-col md:flex-row">
        
        {/* Sidebar Nav */}
        <aside className="w-full md:w-64 bg-white border-r border-neutral-200 flex flex-col shrink-0">
          <div className="p-6 border-b border-neutral-100">
            <div className="flex items-center gap-2 mb-1">
              <div className="p-1.5 bg-blue-600 rounded-lg shadow-sm">
                <CheckCircle2 className="w-5 h-5 text-white" />
              </div>
              <h1 className="font-bold text-lg tracking-tight text-neutral-900 leading-tight">
                AutoComment
              </h1>
            </div>
            <p className="text-[10px] uppercase tracking-widest font-bold text-neutral-400 mt-2">
              Assessment Tool
            </p>
          </div>

          <nav className="flex-1 p-4 space-y-1">
            <button 
              onClick={() => setActiveTab('generator')}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 group ${activeTab === 'generator' ? 'bg-blue-50 text-blue-700 shadow-sm' : 'text-neutral-500 hover:bg-neutral-50 hover:text-neutral-900'}`}
            >
              <LayoutDashboard className={`w-5 h-5 ${activeTab === 'generator' ? 'text-blue-600' : 'text-neutral-400 group-hover:text-neutral-600'}`} />
              <span className="font-medium text-sm">Bảng Điều Khiển</span>
              {activeTab === 'generator' && <ChevronRight className="ml-auto w-4 h-4 opacity-50" />}
            </button>
            <button 
              onClick={() => setActiveTab('settings')}
              className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200 group ${activeTab === 'settings' ? 'bg-blue-50 text-blue-700 shadow-sm' : 'text-neutral-500 hover:bg-neutral-50 hover:text-neutral-900'}`}
            >
              <Database className={`w-5 h-5 ${activeTab === 'settings' ? 'text-blue-600' : 'text-neutral-400 group-hover:text-neutral-600'}`} />
              <span className="font-medium text-sm">Kho Nhận Xét</span>
              {activeTab === 'settings' && <ChevronRight className="ml-auto w-4 h-4 opacity-50" />}
            </button>
          </nav>

          <div className="p-6 border-t border-neutral-100">

            <div className="flex flex-col items-center">
              <div className="w-1.5 h-1.5 rounded-full bg-blue-600 mb-2" />
              <p className="text-[9px] font-black text-neutral-300 uppercase tracking-[0.2em] mb-1">Developed By</p>
              <p className="text-[11px] font-bold text-neutral-800">Lục Minh Sơn</p>
            </div>
          </div>
        </aside>

        {/* Main Content Area */}
        <main className="flex-1 overflow-hidden flex flex-col">
          <header className="h-20 bg-white border-b border-neutral-200 px-8 flex items-center justify-between shrink-0">
            <div>
              <h2 className="text-sm font-semibold text-neutral-500 uppercase tracking-wider">
                {activeTab === 'generator' ? 'Tiến trình làm việc' : 'Cấu hình dữ liệu'}
              </h2>
              <h3 className="text-xl font-bold text-neutral-900">
                {activeTab === 'generator' ? 'Tạo Nhận Xét Tự Động' : 'Quản lý Ngân hàng Câu'}
              </h3>
            </div>
          </header>

          <AnimatePresence mode="wait">
            <motion.div 
              key={activeTab}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
              className="flex-1 overflow-y-auto p-8"
            >
              {activeTab === 'generator' ? (
                <div className="max-w-6xl mx-auto space-y-8">
                  {/* Upload Section */}
                  <section className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
                    <div className="p-8">
                      <div className="flex items-center gap-4 mb-6">
                        <div className="w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center">
                          <Upload className="w-5 h-5 text-blue-600" />
                        </div>
                        <div>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                        <div className="space-y-4">
                          <div className="flex flex-col gap-3">
                            <div 
                              onClick={() => setGradeBlock('k12')}
                              className={`cursor-pointer p-4 rounded-2xl border-2 transition-all duration-200 ${gradeBlock === 'k12' ? 'border-blue-600 bg-blue-50/50' : 'border-neutral-100 hover:border-neutral-200'}`}
                            >
                              <div className="flex items-center gap-3 mb-1">
                                <div className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${gradeBlock === 'k12' ? 'border-blue-600' : 'border-neutral-300'}`}>
                                  {gradeBlock === 'k12' && <div className="w-2 h-2 bg-blue-600 rounded-full" />}
                                </div>
                                <span className="font-bold text-neutral-900 text-sm">Khối 1 & 2</span>
                              </div>
                            </div>

                            <div 
                              onClick={() => setGradeBlock('k345')}
                              className={`cursor-pointer p-4 rounded-2xl border-2 transition-all duration-200 ${gradeBlock === 'k345' ? 'border-blue-600 bg-blue-50/50' : 'border-neutral-100 hover:border-neutral-200'}`}
                            >
                              <div className="flex items-center gap-3 mb-1">
                                <div className={`w-4 h-4 rounded-full border-2 flex items-center justify-center ${gradeBlock === 'k345' ? 'border-blue-600' : 'border-neutral-300'}`}>
                                  {gradeBlock === 'k345' && <div className="w-2 h-2 bg-blue-600 rounded-full" />}
                                </div>
                                <span className="font-bold text-neutral-900 text-sm">Khối 3, 4 & 5</span>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="space-y-4">
                          <div className="flex flex-col gap-3">
                            <div className="grid grid-cols-2 gap-3">
                              <button 
                                onClick={() => setSemester('Học kỳ 1')}
                                className={`flex items-center gap-3 p-3 rounded-xl border-2 transition-all duration-200 ${semester === 'Học kỳ 1' ? 'border-blue-600 bg-blue-50/50' : 'border-neutral-100 hover:border-neutral-200'}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded-full border-2 flex items-center justify-center ${semester === 'Học kỳ 1' ? 'border-blue-600' : 'border-neutral-300'}`}>
                                  {semester === 'Học kỳ 1' && <div className="w-2 h-2 bg-blue-600 rounded-full" />}
                                </div>
                                <span className="font-bold text-neutral-900 text-xs">Học kỳ 1</span>
                              </button>

                              <button 
                                onClick={() => setSemester('Học kỳ 2')}
                                className={`flex items-center gap-3 p-3 rounded-xl border-2 transition-all duration-200 ${semester === 'Học kỳ 2' ? 'border-blue-600 bg-blue-50/50' : 'border-neutral-100 hover:border-neutral-200'}`}
                              >
                                <div className={`w-3.5 h-3.5 rounded-full border-2 flex items-center justify-center ${semester === 'Học kỳ 2' ? 'border-blue-600' : 'border-neutral-300'}`}>
                                  {semester === 'Học kỳ 2' && <div className="w-2 h-2 bg-blue-600 rounded-full" />}
                                </div>
                                <span className="font-bold text-neutral-900 text-xs">Học kỳ 2</span>
                              </button>
                            </div>

                            <div className={`p-3 rounded-xl border-2 transition-all duration-200 ${isCustomSemester ? 'border-blue-600 bg-blue-50/50' : 'border-neutral-100 hover:border-neutral-200'}`}>
                              <div className="flex items-center gap-3 mb-2">
                                <button 
                                  onClick={() => {
                                    setIsCustomSemester(true);
                                    if (customSemester) setSemester(customSemester);
                                  }}
                                  className={`w-3.5 h-3.5 rounded-full border-2 flex items-center justify-center ${isCustomSemester ? 'border-blue-600' : 'border-neutral-300'}`}
                                >
                                  {isCustomSemester && <div className="w-2 h-2 bg-blue-600 rounded-full" />}
                                </button>
                                <span className="font-bold text-neutral-900 text-[11px] uppercase tracking-wider">Học kỳ khác</span>
                              </div>
                              <input 
                                type="text"
                                value={customSemester}
                                placeholder="VD: Học kỳ hè, Cả năm..."
                                onChange={(e) => {
                                  setCustomSemester(e.target.value);
                                  setIsCustomSemester(true);
                                  setSemester(e.target.value || 'Tùy chọn');
                                }}
                                onFocus={() => setIsCustomSemester(true)}
                                className="w-full px-3 py-2 text-xs rounded-lg border border-neutral-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none transition-all bg-white"
                              />
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="flex items-center gap-4">
                        <input 
                          type="file" 
                          accept=".xlsx, .xls, .csv" 
                          ref={fileInputRef}
                          className="hidden"
                          onChange={handleFileUpload}
                        />
                        <button 
                          onClick={() => fileInputRef.current?.click()}
                          className="bg-neutral-900 hover:bg-black text-white px-8 py-3.5 rounded-xl font-bold text-sm transition-all shadow-lg flex items-center gap-2"
                        >
                          <Upload className="w-5 h-5" />
                          Chọn File Excel
                        </button>
                        {fileName && (
                          <div className="flex items-center gap-2 px-4 py-2 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-100 animate-in fade-in slide-in-from-left-2">
                            <CheckCircle2 className="w-4 h-4" />
                            <span className="text-xs font-semibold truncate max-w-[200px]">{fileName}</span>
                          </div>
                        )}
                      </div>

                      {statusMsg.text && (
                        <div className={`mt-6 p-4 rounded-xl text-sm font-medium flex items-start gap-3 animate-in fade-in ease-out duration-300 ${statusMsg.type === 'success' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : statusMsg.type === 'info' ? 'bg-blue-50 text-blue-700 border border-blue-100' : 'bg-rose-50 text-rose-700 border border-rose-100'}`}>
                          {statusMsg.type === 'success' ? <CheckCircle2 className="w-5 h-5 mt-0.5 shrink-0" /> : <AlertCircle className="w-5 h-5 mt-0.5 shrink-0" />}
                          <div className="flex-1">
                            <p>{statusMsg.text}</p>
                          </div>
                        </div>
                      )}
                    </div>
                  </section>

                  {/* Results Table Section */}
                  {students.length > 0 && (
                    <section className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden flex flex-col max-h-[800px]">
                      <div className="p-6 border-b border-neutral-100 flex items-center justify-between bg-neutral-50/50">
                        <h4 className="font-bold text-lg text-neutral-900 flex items-center gap-3">
                          <Users className="w-5 h-5 text-blue-600" />
                          Kết quả xem trước ({students.length} học sinh)
                        </h4>
                        <button 
                          onClick={exportToStandardTemplateExcel}
                          className="bg-green-600 hover:bg-green-700 text-white px-5 py-2.5 rounded-xl text-sm font-semibold flex items-center gap-2 transition-all shadow-[0_4px_12px_rgba(22,163,74,0.3)] hover:translate-y-[-1px] active:translate-y-[0px]"
                        >
                          <FileDown className="w-5 h-5" />
                          Xuất Excel
                        </button>
                      </div>
                      
                      <div className="overflow-x-auto flex-1">
                        <table className="w-full text-left border-collapse min-w-[2000px]">
                          <thead className="bg-neutral-50/80 sticky top-0 z-20 backdrop-blur-sm border-b border-neutral-200">
                            <tr className="text-[10px] uppercase tracking-wider font-extrabold text-neutral-500 bg-neutral-100">
                              <th rowSpan={2} className="px-4 py-3 w-12 text-center border-r border-neutral-200">STT</th>
                              <th rowSpan={2} className="px-4 py-3 w-16 text-center border-r border-neutral-200">Sửa</th>
                              <th rowSpan={2} className="px-4 py-3 w-40 border-r border-neutral-200">Mã học sinh</th>
                              <th rowSpan={2} className="px-4 py-3 w-48 border-r border-neutral-200">Họ và tên</th>
                              <th rowSpan={2} className="px-4 py-3 w-24 text-center border-r border-neutral-200">Học kỳ</th>
                              <th colSpan={4} className="px-4 py-2 text-center border-r border-neutral-200 bg-blue-100/50 text-blue-700">Năng lực chung</th>
                              <th colSpan={gradeBlock === 'k12' ? 6 : 8} className="px-4 py-2 text-center border-r border-neutral-200 bg-emerald-100/50 text-emerald-700">Năng lực đặc thù</th>
                              <th colSpan={6} className="px-4 py-2 text-center bg-rose-100/50 text-rose-700">Phẩm chất</th>
                            </tr>
                            <tr className="text-[9px] uppercase tracking-wider font-bold text-neutral-400">
                              <th className="px-3 py-2 border-r border-neutral-200 bg-blue-50/30">Nhận xét chung</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-blue-50/30">Tự chủ & Tự học</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-blue-50/30">Giao tiếp & Hợp tác</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-blue-50/30">GQVĐ & Sáng tạo</th>
                              
                              <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Nhận xét đặc thù</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Ngôn ngữ</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Tính toán</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Khoa học</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Thẩm mĩ</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Thể chất</th>
                              {gradeBlock === 'k345' && (
                                <>
                                  <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Công nghệ</th>
                                  <th className="px-3 py-2 border-r border-neutral-200 bg-emerald-50/30">Tin học</th>
                                </>
                              )}
                              
                              <th className="px-3 py-2 border-r border-neutral-200 bg-rose-50/30">Nhận xét phẩm chất</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-rose-50/30">Yêu nước</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-rose-50/30">Nhân ái</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-rose-50/30">Chăm chỉ</th>
                              <th className="px-3 py-2 border-r border-neutral-200 bg-rose-50/30">Trung thực</th>
                              <th className="px-3 py-2 bg-rose-50/30">Trách nhiệm</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-neutral-100">
                            {students.map((std, index) => (
                              <tr key={std.id} className="hover:bg-neutral-50/80 transition-colors group">
                                <td className="px-4 py-3 text-center border-r border-neutral-100 text-xs text-neutral-400 font-mono italic">{index + 1}</td>
                                <td className="px-4 py-3 text-center border-r border-neutral-100">
                                  <button onClick={() => openEditModal(std)} className="w-8 h-8 rounded-lg bg-white border border-neutral-200 flex items-center justify-center text-neutral-400 hover:text-blue-600 hover:border-blue-200 hover:bg-blue-50 transition-all shadow-sm">
                                    <Pencil className="w-3.5 h-3.5" />
                                  </button>
                                </td>
                                <td className="px-4 py-3 border-r border-neutral-100 italic text-neutral-500 font-mono text-[10px]">
                                  {std.studentId || 'N/A'}
                                </td>
                                <td className="px-4 py-3 border-r border-neutral-100">
                                  <div className="font-bold text-neutral-900 text-xs">{std.name}</div>
                                </td>
                                <td className="px-4 py-3 border-r border-neutral-100 text-center whitespace-nowrap">
                                  <span className="text-[10px] font-bold text-neutral-500 px-2 py-0.5 bg-neutral-100 rounded">{semester}</span>
                                </td>
                                
                                {/* Năng lực chung columns */}
                                <td className="px-3 py-2 border-r border-neutral-100 bg-blue-50/5 w-64 min-w-[250px]">
                                  <div className="text-[10px] text-neutral-600 italic whitespace-normal">{std.comments.chung}</div>
                                </td>
                                {['Tự chủ và tự học', 'Giao tiếp và hợp tác', 'Giải quyết vấn đề'].map(c => (
                                  <td key={c} className="px-3 py-2 border-r border-neutral-100 bg-blue-50/5 w-48 min-w-[150px]">
                                    <div className="text-[10px] text-neutral-700 leading-tight mb-1">{std.individualComments[c]}</div>
                                    <div className="text-[9px] font-bold text-blue-600">{std.originalGrades[c]}</div>
                                  </td>
                                ))}

                                {/* Năng lực đặc thù columns */}
                                <td className="px-3 py-2 border-r border-neutral-100 bg-emerald-50/5 w-64 min-w-[250px]">
                                  <div className="text-[10px] text-neutral-600 italic whitespace-normal">{std.comments.dacThu}</div>
                                </td>
                                {(['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Thẩm mĩ', 'Thể chất', ...(gradeBlock === 'k345' ? ['Công nghệ', 'Tin học'] : [])] as string[]).map(c => (
                                  <td key={c} className="px-3 py-2 border-r border-neutral-100 bg-emerald-50/5 w-48 min-w-[150px]">
                                    <div className="text-[10px] text-neutral-700 leading-tight mb-1">{std.individualComments[c]}</div>
                                    <div className="text-[9px] font-bold text-emerald-600">{std.originalGrades[c]}</div>
                                  </td>
                                ))}

                                {/* Phẩm chất columns */}
                                <td className="px-3 py-2 border-r border-neutral-100 bg-rose-50/5 w-64 min-w-[250px]">
                                  <div className="text-[10px] text-neutral-600 italic whitespace-normal">{std.comments.phamChat}</div>
                                </td>
                                {['Yêu nước', 'Nhân ái', 'Chăm chỉ', 'Trung thực', 'Trách nhiệm'].map(c => (
                                  <td key={c} className="px-3 py-2 border-r border-neutral-200 bg-rose-50/5 w-48 min-w-[150px]">
                                    <div className="text-[10px] text-neutral-700 leading-tight mb-1">{std.individualComments[c]}</div>
                                    <div className="text-[9px] font-bold text-rose-600">{std.originalGrades[c]}</div>
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </section>
                  )}
                </div>
              ) : (
                <div className="max-w-4xl mx-auto space-y-6">
                  <div className="bg-blue-600 rounded-3xl p-8 text-white relative overflow-hidden shadow-2xl">
                    <div className="relative z-10">
                      <div className="flex items-center gap-3 mb-4">
                        <Database className="w-8 h-8 text-blue-200" />
                        <h4 className="text-2xl font-bold">Cấu hình Ngân hàng Câu</h4>
                      </div>

                    </div>
                    <div className="absolute top-[-10%] right-[-5%] w-64 h-64 bg-white/10 rounded-full blur-3xl" />
                  </div>
                  
                  <div className="bg-white rounded-2xl border border-neutral-200 shadow-xl overflow-hidden">
                    <div className="p-4 bg-neutral-50 border-b border-neutral-200 flex flex-col sm:flex-row items-center justify-between gap-4">
                      <div className="flex flex-wrap items-center gap-3">
                        <span className="flex items-center gap-2 text-xs font-bold text-neutral-400 uppercase tracking-widest mr-2">
                          <Settings className="w-3 h-3" /> Chỉnh sửa
                        </span>
                        <button 
                          onClick={saveCommentBank}
                          className="flex items-center gap-2 px-5 py-2 bg-blue-600 text-white rounded-xl text-[11px] font-bold hover:bg-blue-700 transition-all shadow-[0_4px_12px_rgba(37,99,235,0.2)] hover:translate-y-[-1px] active:translate-y-[0px]"
                        >
                          <Save className="w-3.5 h-3.5" />
                          Lưu Thay Đổi
                        </button>
                        <button 
                          onClick={importCommentBank}
                          className="flex items-center gap-2 px-5 py-2 bg-neutral-100 text-neutral-600 rounded-xl text-[11px] font-bold hover:bg-neutral-200 transition-all"
                        >
                          <Upload className="w-3.5 h-3.5" />
                          Nạp File (CSV)
                        </button>
                        <button 
                          onClick={downloadCommentBank}
                          className="flex items-center gap-2 px-5 py-2 bg-neutral-100 text-neutral-600 rounded-xl text-[11px] font-bold hover:bg-neutral-200 transition-all"
                        >
                          <FileDown className="w-3.5 h-3.5" />
                          Tải Xuống
                        </button>
                        <button 
                          onClick={resetCommentBank}
                          className="flex items-center gap-2 px-5 py-2 bg-white border border-neutral-200 text-rose-500 rounded-xl text-[11px] font-bold hover:bg-rose-50 hover:border-rose-100 transition-all"
                        >
                          Phục Hồi Mặc Định
                        </button>
                      </div>

                    </div>
                    {statusMsg.text && activeTab === 'settings' && (
                       <div className={`m-4 p-3 rounded-xl text-[11px] font-bold flex items-center gap-2 animate-in fade-in slide-in-from-top-2 ${statusMsg.type === 'success' ? 'bg-emerald-50 text-emerald-700' : 'bg-blue-50 text-blue-700'}`}>
                         {statusMsg.type === 'success' ? <CheckCircle2 className="w-3.5 h-3.5" /> : <Info className="w-3.5 h-3.5" />}
                         {statusMsg.text}
                       </div>
                    )}
                    <textarea
                      value={commentBankText}
                      onChange={(e) => setCommentBankText(e.target.value)}
                      className="w-full h-[600px] p-8 font-mono text-sm leading-relaxed focus:outline-none bg-neutral-50/30 text-neutral-700 selection:bg-blue-100"
                      spellCheck="true"
                    />
                  </div>
                </div>
              )}
            </motion.div>
          </AnimatePresence>
        </main>
      </div>

      {/* Edit Modal (Portal-like) */}
      <AnimatePresence>
        {editingStudent && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={closeEditModal}
              className="absolute inset-0 bg-neutral-900/40 backdrop-blur-md" 
            />
            
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative w-full max-w-6xl bg-white rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
            >
              {/* Modal Header */}
              <header className="px-8 py-6 bg-neutral-900 text-white flex justify-between items-center shrink-0">
                <div className="flex items-baseline gap-4">
                  <h2 className="text-2xl font-black tracking-tight">Hiệu chỉnh chi tiết</h2>
                  <span className="text-white/40 text-sm font-medium tracking-widest uppercase">Student Profile</span>
                </div>
                <div className="flex items-center gap-4">
                  <div className="text-right hidden sm:block">
                     <p className="text-lg font-bold text-blue-400 leading-tight">{editingStudent.name}</p>
                     <p className="text-[10px] font-bold text-white/50 tracking-widest uppercase">{editingStudent.studentId || 'No ID'}</p>
                  </div>
                  <button onClick={closeEditModal} className="w-10 h-10 flex items-center justify-center rounded-xl bg-white/10 hover:bg-white/20 transition-colors">
                    <X className="w-6 h-6" />
                  </button>
                </div>
              </header>

              {/* Scroll Content */}
              <div className="flex-1 overflow-y-auto p-8 bg-neutral-50 space-y-12">
                
                {/* 1. Năng lực chung */}
                <div className="space-y-6">
                  <div className="flex items-center justify-between">
                    <h5 className="text-sm font-black uppercase tracking-widest text-neutral-400">01. Năng lực chung</h5>
                    <button onClick={regenerateGroupedForEditingStudent} className="text-[10px] font-bold text-blue-600 bg-blue-50 px-3 py-1.5 rounded-full hover:bg-blue-100 transition-colors">Làm mới nội dung gộp</button>
                  </div>
                  <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
                    <div className="lg:col-span-1 space-y-2">
                       <label className="text-[11px] font-black uppercase text-neutral-500">Nội dung gộp</label>
                       <textarea 
                         value={editingStudent.comments.chung}
                         onChange={(e) => handleEditGroupedComment('chung', e.target.value)}
                         className="w-full h-32 p-4 text-xs font-semibold rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-blue-500 outline-none resize-none leading-relaxed italic"
                         spellCheck="true"
                       />
                    </div>
                    <div className="lg:col-span-3 grid grid-cols-1 md:grid-cols-3 gap-6">
                      {['Tự chủ và tự học', 'Giao tiếp và hợp tác', 'Giải quyết vấn đề'].map(crit => (
                        <div key={crit} className="space-y-2">
                           <div className="flex justify-between items-center">
                             <label className="text-[10px] font-black uppercase text-neutral-500">{crit}</label>
                             <span className="text-[9px] font-black bg-blue-600 text-white px-2 py-0.5 rounded-full">{editingStudent.originalGrades[crit] || 'N/A'}</span>
                           </div>
                           <textarea 
                             value={editingStudent.individualComments[crit] || ''}
                             onChange={(e) => handleEditIndividualComment(crit, e.target.value)}
                             className="w-full h-32 p-4 text-xs font-medium rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-blue-500 outline-none resize-none leading-relaxed"
                             spellCheck="true"
                           />
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* 2. Năng lực đặc thù */}
                <div className="space-y-6">
                  <h5 className="text-sm font-black uppercase tracking-widest text-neutral-400">02. Năng lực đặc thù</h5>
                  <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
                    <div className="lg:col-span-1 space-y-2">
                       <label className="text-[11px] font-black uppercase text-neutral-500">Nội dung gộp</label>
                       <textarea 
                         value={editingStudent.comments.dacThu}
                         onChange={(e) => handleEditGroupedComment('dacThu', e.target.value)}
                         className="w-full h-32 p-4 text-xs font-semibold rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-emerald-500 outline-none resize-none leading-relaxed italic"
                         spellCheck="true"
                       />
                    </div>
                    <div className="lg:col-span-3 grid grid-cols-2 md:grid-cols-4 gap-6">
                      {['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Thẩm mĩ', 'Thể chất'].map(crit => (
                        <div key={crit} className="space-y-2">
                           <div className="flex justify-between items-center">
                             <label className="text-[10px] font-black uppercase text-neutral-500">{crit}</label>
                             <span className="text-[9px] font-black bg-emerald-600 text-white px-2 py-0.5 rounded-full">{editingStudent.originalGrades[crit] || 'N/A'}</span>
                           </div>
                           <textarea 
                             value={editingStudent.individualComments[crit] || ''}
                             onChange={(e) => handleEditIndividualComment(crit, e.target.value)}
                             className="w-full h-32 p-4 text-[10px] font-medium rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-emerald-500 outline-none resize-none leading-relaxed"
                             spellCheck="true"
                           />
                        </div>
                      ))}
                      {gradeBlock === 'k345' && ['Công nghệ', 'Tin học'].map(crit => (
                        <div key={crit} className="space-y-2">
                           <div className="flex justify-between items-center">
                             <label className="text-[10px] font-black uppercase text-neutral-500">{crit}</label>
                             <span className="text-[9px] font-black bg-blue-500 text-white px-2 py-0.5 rounded-full">{editingStudent.originalGrades[crit] || 'N/A'}</span>
                           </div>
                           <textarea 
                             value={editingStudent.individualComments[crit] || ''}
                             onChange={(e) => handleEditIndividualComment(crit, e.target.value)}
                             className="w-full h-32 p-4 text-[10px] font-medium rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-blue-500 outline-none resize-none leading-relaxed"
                             spellCheck="true"
                           />
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* 3. Phẩm chất */}
                <div className="space-y-6">
                  <h5 className="text-sm font-black uppercase tracking-widest text-neutral-400">03. Phẩm chất</h5>
                  <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
                    <div className="lg:col-span-1 space-y-2">
                       <label className="text-[11px] font-black uppercase text-neutral-500">Nội dung gộp</label>
                       <textarea 
                         value={editingStudent.comments.phamChat}
                         onChange={(e) => handleEditGroupedComment('phamChat', e.target.value)}
                         className="w-full h-32 p-4 text-xs font-semibold rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-rose-500 outline-none resize-none leading-relaxed italic"
                         spellCheck="true"
                       />
                    </div>
                    <div className="lg:col-span-3 grid grid-cols-2 md:grid-cols-5 gap-6">
                      {['Yêu nước', 'Nhân ái', 'Chăm chỉ', 'Trung thực', 'Trách nhiệm'].map(crit => (
                        <div key={crit} className="space-y-2">
                           <div className="flex justify-between items-center">
                             <label className="text-[10px] font-black uppercase text-neutral-500">{crit}</label>
                             <span className="text-[9px] font-black bg-rose-600 text-white px-2 py-0.5 rounded-full">{editingStudent.originalGrades[crit] || 'N/A'}</span>
                           </div>
                           <textarea 
                             value={editingStudent.individualComments[crit] || ''}
                             onChange={(e) => handleEditIndividualComment(crit, e.target.value)}
                             className="w-full h-32 p-4 text-[10px] font-medium rounded-2xl bg-white border border-neutral-200 focus:ring-2 focus:ring-rose-500 outline-none resize-none leading-relaxed"
                             spellCheck="true"
                           />
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

              </div>

              {/* Modal Footer */}
              <footer className="px-8 py-6 border-t border-neutral-200 flex justify-end gap-3 shrink-0 bg-white shadow-[0_-10px_20px_rgba(0,0,0,0.02)]">
                <button 
                  onClick={closeEditModal}
                  className="px-8 py-3 rounded-xl font-bold text-sm text-neutral-500 hover:text-neutral-900 transition-colors"
                >
                  Hủy bỏ
                </button>
                <button 
                  onClick={saveEditedStudent}
                  className="px-10 py-3.5 bg-neutral-900 hover:bg-black text-white rounded-xl font-black text-sm flex items-center gap-2 transition-all shadow-xl hover:translate-y-[-2px]"
                >
                  <Save className="w-5 h-5" />
                  Xác nhận & Lưu
                </button>
              </footer>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
