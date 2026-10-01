/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Calendar, 
  Clock, 
  Users, 
  Phone, 
  User, 
  FileSpreadsheet, 
  RotateCcw, 
  Search, 
  CheckCircle2, 
  Sparkles,
  ChevronDown,
  Building2,
  FileText
} from 'lucide-react';

interface Reservation {
  id: string;
  createdAt: string;
  name: string;
  phone: string;
  date: string;
  time: string;
  attendees: number;
  room: string;
  notes: string;
}

const STORAGE_KEY = 'study_room_reservations_vite_v1';

// 1시간 간격 드롭다운 옵션 생성 (00:00~01:00부터 23:00~24:00까지)
const TIME_SLOTS: string[] = Array.from({ length: 24 }, (_, i) => {
  const start = String(i).padStart(2, '0') + ':00';
  const end = String(i + 1).padStart(2, '0') + ':00';
  return `${start}~${end}`;
});

const STUDY_ROOMS = [
  '스터디 1실 (대형 / 빔프로젝터 완비)',
  '스터디 2실 (집중형 / 대형 모니터 완비)',
  '스터디 3실 (토론형 / 와이드 화이트보드)',
];

// 성만 노출되도록 마스킹 처리 (예: 홍길동 -> 홍**, 김철수 -> 김**, 박준 -> 박*)
function maskName(name: string): string {
  if (!name) return '';
  const trimmed = name.trim();
  if (trimmed.length <= 1) return trimmed + '*';
  if (trimmed.length === 2) return trimmed[0] + '*';
  return trimmed[0] + '*'.repeat(trimmed.length - 1);
}

// 전화번호 자동 하이픈 포맷팅
function formatPhoneNumber(value: string): string {
  const clean = value.replace(/[^0-9]/g, '');
  if (clean.length < 4) return clean;
  if (clean.length < 7) return clean.replace(/(\d{3})(\d{1,3})/, '$1-$2');
  if (clean.length < 11) return clean.replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');
  return clean.slice(0, 11).replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
}

export default function App() {
  const getTodayString = () => {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };

  const todayStr = getTodayString();

  // 폼 입력 상태
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [date, setDate] = useState(todayStr);
  const [time, setTime] = useState('14:00~15:00');
  const [attendees, setAttendees] = useState<number>(1);
  const [room, setRoom] = useState(STUDY_ROOMS[0]);
  const [notes, setNotes] = useState('');

  // 검색 및 데이터 목록
  const [searchQuery, setSearchQuery] = useState('');
  const [reservations, setReservations] = useState<Reservation[]>([]);
  const [toastMessage, setToastMessage] = useState<{ title: string; desc: string } | null>(null);

  // 로컬 스토리지 초기화 및 로드
  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        setReservations(JSON.parse(saved));
        return;
      } catch (e) {
        console.error('Failed to parse reservations from localStorage', e);
      }
    }

    // 기본 예시 데이터
    const initialData: Reservation[] = [
      {
        id: 'RES-1001',
        createdAt: '2026-09-30 14:20',
        name: '김서연',
        phone: '010-3849-1029',
        date: todayStr,
        time: '14:00~15:00',
        attendees: 4,
        room: STUDY_ROOMS[0],
        notes: '빔프로젝터 및 화이트보드 사용 예정',
      },
      {
        id: 'RES-1002',
        createdAt: '2026-09-30 15:45',
        name: '이도현',
        phone: '010-9948-2811',
        date: todayStr,
        time: '18:00~19:00',
        attendees: 2,
        room: STUDY_ROOMS[1],
        notes: '듀얼모니터 연결 작업',
      },
      {
        id: 'RES-1003',
        createdAt: '2026-09-30 16:10',
        name: '박지민',
        phone: '010-5542-1920',
        date: todayStr,
        time: '20:00~21:00',
        attendees: 3,
        room: STUDY_ROOMS[2],
        notes: '독서 스터디 모임',
      },
    ];

    setReservations(initialData);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(initialData));
  }, []);

  // 토스트 자동 사라짐
  useEffect(() => {
    if (toastMessage) {
      const timer = setTimeout(() => setToastMessage(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [toastMessage]);

  // 엑셀(.xlsx) 파일 생성 및 다운로드 함수
  const exportExcel = (data: Reservation[], filenamePrefix = '스터디실_대여신청내역') => {
    if (data.length === 0) {
      alert('저장할 예약 신청 내역이 없습니다.');
      return;
    }

    const now = new Date();
    const timeStamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}`;
    const filename = `${filenamePrefix}_${timeStamp}.xlsx`;

    const headers = [
      '신청 번호',
      '신청 접수 일시',
      '신청자 성명',
      '전화번호',
      '예약 일자',
      '예약 이용 시간',
      '이용 인원수(명)',
      '배정 스터디실',
      '요청사항 및 메모',
    ];

    const rows = data.map((item, index) => [
      item.id || `RES-${1000 + index + 1}`,
      item.createdAt || '-',
      item.name,
      item.phone,
      item.date,
      item.time,
      item.attendees,
      item.room,
      item.notes || '없음',
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    // 컬럼 너비 지정
    worksheet['!cols'] = [
      { wch: 14 },
      { wch: 18 },
      { wch: 14 },
      { wch: 16 },
      { wch: 14 },
      { wch: 18 },
      { wch: 14 },
      { wch: 28 },
      { wch: 30 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '스터디실신청명단');
    XLSX.writeFile(workbook, filename);
  };

  // 폼 제출 핸들러 (유효성 검사 -> 로컬스토리지 저장 -> 엑셀 자동 다운로드)
  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    if (!name.trim()) {
      alert('신청자 이름을 입력해주세요.');
      return;
    }
    if (!phone.trim() || phone.replace(/[^0-9]/g, '').length < 10) {
      alert('올바른 휴대폰 번호를 입력해주세요. (예: 010-1234-5678)');
      return;
    }
    if (!date) {
      alert('예약 날짜를 선택해주세요.');
      return;
    }
    if (!time) {
      alert('예약 시간을 선택해주세요.');
      return;
    }
    if (attendees < 1 || attendees > 7) {
      alert('인원수는 최소 1명에서 최대 7명까지 가능합니다.');
      return;
    }

    const now = new Date();
    const createdAt = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const newId = `RES-${Date.now().toString().slice(-6)}`;

    const newEntry: Reservation = {
      id: newId,
      createdAt,
      name: name.trim(),
      phone: phone.trim(),
      date,
      time,
      attendees,
      room,
      notes: notes.trim(),
    };

    const updated = [newEntry, ...reservations];
    setReservations(updated);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));

    // 제출하기 클릭 시 즉시 엑셀 파일 다운로드
    exportExcel(updated, `스터디실_대여신청_신규접수_${newEntry.name}`);

    // 성공 토스트
    setToastMessage({
      title: '대여 신청 완료 🎉',
      desc: `[${newEntry.name}님] ${newEntry.date} (${newEntry.time}) 예약이 접수되었으며 엑셀 파일이 다운로드되었습니다.`,
    });

    // 폼 초기화
    handleReset();
  };

  // 폼 리셋 핸들러
  const handleReset = () => {
    setName('');
    setPhone('');
    setDate(todayStr);
    setTime('14:00~15:00');
    setAttendees(1);
    setRoom(STUDY_ROOMS[0]);
    setNotes('');
  };

  // 필터링된 예약 목록 (신청자 마스킹 검색 또는 일자 검색 지원)
  const filteredReservations = reservations.filter((r) => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return true;
    return (
      r.name.toLowerCase().includes(q) ||
      maskName(r.name).toLowerCase().includes(q) ||
      r.date.includes(q) ||
      r.time.includes(q) ||
      r.room.toLowerCase().includes(q)
    );
  });

  const todayCount = reservations.filter((r) => r.date === todayStr).length;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 py-10 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-4xl mx-auto">
        
        {/* 상단: 스터디실 대여📚 문구 중앙에 크게 배치 */}
        <header className="text-center mb-10">
          <h1 className="text-4xl sm:text-5xl font-extrabold text-slate-900 tracking-tight flex items-center justify-center gap-3">
            스터디실 대여📚
          </h1>
          <p className="mt-3 text-base sm:text-lg text-slate-600 font-normal">
            쾌적한 학습과 몰입을 위한 공간을 손쉽게 예약하고 관리하세요
          </p>
        </header>

        {/* 중앙 정렬된 모던 화이트 톤 입력 폼 */}
        <main className="space-y-8">
          <section className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6 sm:p-10 transition-all hover:shadow-md">
            <div className="border-b border-slate-100 pb-5 mb-6">
              <h2 className="text-xl font-bold text-slate-900 flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-blue-600" />
                대여 신청서 작성
              </h2>
              <p className="text-sm text-slate-500 mt-1">
                신청 정보를 입력하고 제출하시면 즉시 엑셀(.xlsx) 파일로 저장됩니다.
              </p>
            </div>

            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                
                {/* 1. 신청자 이름 */}
                <div>
                  <label htmlFor="name" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <User className="w-4 h-4 text-slate-500" />
                    신청자 이름 <span className="text-rose-500 font-bold">*</span>
                  </label>
                  <input
                    id="name"
                    type="text"
                    required
                    value={name}
                    onChange={(e) => setName(e.target.value)}
                    placeholder="예: 홍길동"
                    className="w-full h-11 px-3.5 rounded-lg border border-slate-300 bg-white text-slate-900 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors placeholder:text-slate-400"
                  />
                  <p className="text-xs text-slate-400 mt-1">예약자 본인의 실명을 입력해주세요.</p>
                </div>

                {/* 2. 전화번호 */}
                <div>
                  <label htmlFor="phone" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <Phone className="w-4 h-4 text-slate-500" />
                    전화번호 <span className="text-rose-500 font-bold">*</span>
                  </label>
                  <input
                    id="phone"
                    type="tel"
                    required
                    maxLength={13}
                    value={phone}
                    onChange={(e) => setPhone(formatPhoneNumber(e.target.value))}
                    placeholder="예: 010-1234-5678"
                    className="w-full h-11 px-3.5 rounded-lg border border-slate-300 bg-white text-slate-900 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors placeholder:text-slate-400"
                  />
                  <p className="text-xs text-slate-400 mt-1">숫자만 입력 시 자동으로 하이픈이 생성됩니다.</p>
                </div>

                {/* 3. 예약 날짜 */}
                <div>
                  <label htmlFor="date" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <Calendar className="w-4 h-4 text-slate-500" />
                    예약 일자 <span className="text-rose-500 font-bold">*</span>
                  </label>
                  <input
                    id="date"
                    type="date"
                    required
                    min={todayStr}
                    value={date}
                    onChange={(e) => setDate(e.target.value)}
                    className="w-full h-11 px-3.5 rounded-lg border border-slate-300 bg-white text-slate-900 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
                  />
                  <p className="text-xs text-slate-400 mt-1">오늘 이후의 날짜만 선택 가능합니다.</p>
                </div>

                {/* 4. 예약 시간 (한시간 간격 드롭다운) */}
                <div>
                  <label htmlFor="time" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <Clock className="w-4 h-4 text-slate-500" />
                    예약 시간 <span className="text-rose-500 font-bold">*</span>
                  </label>
                  <div className="relative">
                    <select
                      id="time"
                      required
                      value={time}
                      onChange={(e) => setTime(e.target.value)}
                      className="w-full h-11 px-3.5 pr-10 rounded-lg border border-slate-300 bg-white text-slate-900 text-sm appearance-none focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors cursor-pointer"
                    >
                      {TIME_SLOTS.map((slot) => (
                        <option key={slot} value={slot}>
                          {slot} (1시간)
                        </option>
                      ))}
                    </select>
                    <ChevronDown className="w-4 h-4 text-slate-400 absolute right-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  </div>
                  <p className="text-xs text-slate-400 mt-1">1시간 단위(ex. 00:00~01:00)로 선택 가능합니다.</p>
                </div>

                {/* 5. 인원수 (number 타입, min: 1, max: 7) */}
                <div>
                  <label htmlFor="attendees" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <Users className="w-4 h-4 text-slate-500" />
                    이용 인원수 <span className="text-rose-500 font-bold">*</span>
                  </label>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setAttendees((prev) => Math.max(1, prev - 1))}
                      className="w-11 h-11 flex items-center justify-center rounded-lg border border-slate-300 bg-slate-50 text-slate-700 hover:bg-slate-100 active:scale-95 transition-all font-semibold text-base"
                    >
                      −
                    </button>
                    <input
                      id="attendees"
                      type="number"
                      min={1}
                      max={7}
                      required
                      value={attendees}
                      onChange={(e) => {
                        const val = parseInt(e.target.value, 10);
                        if (!isNaN(val)) {
                          setAttendees(Math.min(7, Math.max(1, val)));
                        }
                      }}
                      className="flex-1 h-11 px-3.5 rounded-lg border border-slate-300 bg-white text-slate-900 text-center text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
                    />
                    <button
                      type="button"
                      onClick={() => setAttendees((prev) => Math.min(7, prev + 1))}
                      className="w-11 h-11 flex items-center justify-center rounded-lg border border-slate-300 bg-slate-50 text-slate-700 hover:bg-slate-100 active:scale-95 transition-all font-semibold text-base"
                    >
                      +
                    </button>
                  </div>
                  <p className="text-xs text-slate-400 mt-1">최소 1명부터 최대 7명까지 입력 가능합니다.</p>
                </div>

                {/* 6. 스터디실 선택 */}
                <div>
                  <label htmlFor="room" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <Building2 className="w-4 h-4 text-slate-500" />
                    스터디실 선택
                  </label>
                  <div className="relative">
                    <select
                      id="room"
                      value={room}
                      onChange={(e) => setRoom(e.target.value)}
                      className="w-full h-11 px-3.5 pr-10 rounded-lg border border-slate-300 bg-white text-slate-900 text-sm appearance-none focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors cursor-pointer"
                    >
                      {STUDY_ROOMS.map((r) => (
                        <option key={r} value={r}>
                          {r}
                        </option>
                      ))}
                    </select>
                    <ChevronDown className="w-4 h-4 text-slate-400 absolute right-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  </div>
                  <p className="text-xs text-slate-400 mt-1">원하시는 공간의 특성에 맞게 선택하세요.</p>
                </div>

                {/* 7. 기타 요청사항 */}
                <div className="sm:col-span-2">
                  <label htmlFor="notes" className="block text-sm font-semibold text-slate-800 mb-1.5 flex items-center gap-1.5">
                    <FileText className="w-4 h-4 text-slate-500" />
                    기타 요청사항
                  </label>
                  <input
                    id="notes"
                    type="text"
                    value={notes}
                    onChange={(e) => setNotes(e.target.value)}
                    placeholder="예: 빔프로젝터 HDMI 케이블 필요, 추가 멀티탭 구비 희망"
                    className="w-full h-11 px-3.5 rounded-lg border border-slate-300 bg-white text-slate-900 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors placeholder:text-slate-400"
                  />
                </div>
              </div>

              {/* 엑셀 자동 저장 안내 */}
              <div className="bg-slate-50 border border-dashed border-slate-300 rounded-xl p-3.5 flex items-center gap-3 text-xs text-slate-600">
                <FileSpreadsheet className="w-5 h-5 text-emerald-600 shrink-0" />
                <span>
                  <strong>엑셀 자동 저장:</strong> [제출하기] 버튼을 누르면 입력하신 신청 내역이 즉시 <strong>엑셀(.xlsx) 파일</strong>로 자동 다운로드됩니다.
                </span>
              </div>

              {/* 버튼 그룹: 제출하기와 리셋 */}
              <div className="flex flex-col sm:flex-row items-center justify-center gap-3 pt-3 border-t border-slate-100">
                <button
                  type="submit"
                  className="w-full sm:w-auto px-8 h-12 bg-slate-900 hover:bg-slate-800 text-white font-semibold text-sm rounded-xl shadow-sm hover:shadow active:scale-[0.99] transition-all flex items-center justify-center gap-2"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  제출하기 (엑셀 저장)
                </button>
                <button
                  type="button"
                  onClick={handleReset}
                  className="w-full sm:w-auto px-6 h-12 bg-white hover:bg-slate-50 text-slate-600 hover:text-slate-900 border border-slate-300 font-semibold text-sm rounded-xl transition-all flex items-center justify-center gap-2"
                >
                  <RotateCcw className="w-4 h-4" />
                  리셋
                </button>
              </div>
            </form>
          </section>

          {/* 예약 내역 조회 카드 (신청자(마스킹), 일자, 시간, 스터디실만 표시) */}
          <section className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6 sm:p-10 transition-all hover:shadow-md">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-slate-100 pb-5 mb-6">
              <div>
                <h2 className="text-xl font-bold text-slate-900 flex items-center gap-2">
                  <Calendar className="w-5 h-5 text-blue-600" />
                  예약 내역 조회
                </h2>
                <p className="text-sm text-slate-500 mt-1">
                  개인정보 보호를 위해 신청자명은 성만 노출(김**)되며, 예약 일정 정보만 조회됩니다.
                </p>
              </div>

              <button
                type="button"
                onClick={() => exportExcel(reservations)}
                className="inline-flex items-center justify-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold rounded-lg shadow-sm transition-all"
              >
                <FileSpreadsheet className="w-4 h-4" />
                엑셀(.xlsx) 전체 다운로드
              </button>
            </div>

            {/* 통계 요약 카드 */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
              <div className="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                <div className="text-xs text-slate-500 font-medium mb-1">총 신청 건수</div>
                <div className="text-2xl font-bold text-slate-900 font-mono">{reservations.length}건</div>
              </div>
              <div className="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                <div className="text-xs text-slate-500 font-medium mb-1">오늘 예약 건수</div>
                <div className="text-2xl font-bold text-slate-900 font-mono">{todayCount}건</div>
              </div>
              <div className="bg-slate-50 border border-slate-200/80 rounded-xl p-4">
                <div className="text-xs text-slate-500 font-medium mb-1">운영 스터디실</div>
                <div className="text-2xl font-bold text-slate-900 font-mono">3개실</div>
              </div>
            </div>

            {/* 검색창 */}
            <div className="mb-4">
              <div className="relative">
                <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="신청자 또는 예약 일자로 검색..."
                  className="w-full h-10 pl-10 pr-4 rounded-lg border border-slate-200 bg-white text-slate-900 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 transition-colors placeholder:text-slate-400"
                />
              </div>
            </div>

            {/* 테이블 (신청자(김**), 예약 일자, 예약 시간, 스터디실만 표시 / 연락처, 인원, 삭제 제외) */}
            <div className="overflow-x-auto rounded-xl border border-slate-200/80">
              <table className="w-full text-left text-sm border-collapse">
                <thead>
                  <tr className="bg-slate-50/80 text-slate-600 font-semibold border-b border-slate-200">
                    <th className="py-3 px-4 w-14">No.</th>
                    <th className="py-3 px-4">신청자</th>
                    <th className="py-3 px-4">예약 일자</th>
                    <th className="py-3 px-4">예약 시간</th>
                    <th className="py-3 px-4">스터디실</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredReservations.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="py-12 text-center text-slate-400">
                        <Calendar className="w-8 h-8 mx-auto mb-2 text-slate-300" />
                        {searchQuery ? '검색 조건에 맞는 예약 내역이 없습니다.' : '등록된 대여 신청 내역이 없습니다.'}
                      </td>
                    </tr>
                  ) : (
                    filteredReservations.map((item, index) => (
                      <tr key={item.id} className="hover:bg-slate-50/60 transition-colors">
                        <td className="py-3.5 px-4 font-mono text-slate-500 text-xs">{index + 1}</td>
                        <td className="py-3.5 px-4 font-semibold text-slate-900">{maskName(item.name)}</td>
                        <td className="py-3.5 px-4 font-mono text-slate-700 text-xs">{item.date}</td>
                        <td className="py-3.5 px-4 font-medium text-slate-900">{item.time}</td>
                        <td className="py-3.5 px-4">
                          <span className="inline-block px-2.5 py-0.5 text-xs font-medium rounded-md bg-blue-50 text-blue-700 border border-blue-100">
                            {item.room}
                          </span>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </main>

        <footer className="text-center mt-12 text-xs text-slate-400">
          © 2026 스터디실 대여 신청 관리 시스템 · Clean White Edition
        </footer>
      </div>

      {/* 완료 알림 토스트 */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-white border border-slate-200 shadow-xl rounded-xl p-4 max-w-sm flex items-start gap-3 transition-all animate-in fade-in slide-in-from-bottom-2">
          <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
          <div className="flex-1">
            <h4 className="text-sm font-bold text-slate-900">{toastMessage.title}</h4>
            <p className="text-xs text-slate-600 mt-0.5 leading-relaxed">{toastMessage.desc}</p>
          </div>
          <button
            onClick={() => setToastMessage(null)}
            className="text-slate-400 hover:text-slate-600 text-xs p-1"
          >
            ✕
          </button>
        </div>
      )}
    </div>
  );
}
