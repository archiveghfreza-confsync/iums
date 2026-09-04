
//==========================================================
//=======Date Time Functions================================
//==========================================================
function do_date_g2jnow() {
    const g_date = new Date()
    const j_date = new Intl.DateTimeFormat('fa-IR', {
        year: 'numeric',   // سال ۴ رقمی
        month: '2-digit',   // دو رقمی
        day: '2-digit'    // دو رقمی
    }).format(g_date);
    return j_date
}
function do_date_g2j(gy, gm, gd) {
    const g_date = new Date(gy, gm, gd)
    const j_date = new Intl.DateTimeFormat('fa-IR').format(g_date);
    return j_date
}
//==========================================================
//==========================================================
//==========================================================
//==========================================================

/*
  راه‌اندازی تقویم شمسی روی یک فیلد ورودی
  @param {string} elementId - شناسه المان ورودی
  @param {object} options - تنظیمات اختیاری
 */
function initPersianDatepicker(elementId, options = {}) {
    const element = document.getElementById(elementId);
    if (!element) {
        console.warn(`⚠️ عنصر با شناسه "${elementId}" یافت نشد`);
        return false;
    }

    // بررسی وجود کتابخانه‌ها
    if (typeof $ === 'undefined') {
        console.error('❌ jQuery بارگذاری نشده است');
        return false;
    }
    if (typeof $.fn.persianDatepicker === 'undefined') {
        console.error('❌ persian-datepicker بارگذاری نشده است');
        return false;
    }

    try {
        // حذف دیت‌پیکر قبلی (در صورت وجود)
        if ($(element).data('persianDatepicker')) {
            $(element).data('persianDatepicker').destroy();
        }

        // تنظیمات پیش‌فرض
        const defaultOptions = {
            format: 'YYYY/MM/DD',
            autoClose: true,
            initialValue: true,
            initialValueType: 'persian',
            initialValue: new persianDate().format('YYYY/MM/DD'),
            onSelect: function (unixDate) {
                const formatted = new persianDate(unixDate).format('YYYY/MM/DD');
                element.value = formatted;

                // رویداد سفارشی برای اطلاع از تغییر تاریخ
                element.dispatchEvent(new CustomEvent('dateSelected', {
                    detail: {
                        unixDate: unixDate,
                        formattedDate: formatted
                    }
                }));
            }
        };

        // ترکیب تنظیمات
        const settings = { ...defaultOptions, ...options };

        // راه‌اندازی دیت‌پیکر
        $(element).persianDatepicker(settings);

        // تنظیم مقدار اولیه (اگر خالی باشد)
        if (!element.value) {
            element.value = new persianDate().format('YYYY/MM/DD');
        }

        console.log(`✅ تقویم روی "${elementId}" با موفقیت راه‌اندازی شد`);
        return true;

    } catch (error) {
        console.error(`❌ خطا در راه‌اندازی تقویم "${elementId}":`, error);
        return false;
    }
}

/**
 * راه‌اندازی تقویم روی چند المان همزمان
 * @param {array} elementIds - آرایه‌ای از شناسه‌ها
 */
function initMultipleDatepickers(elementIds) {
    if (!Array.isArray(elementIds)) {
        console.warn('⚠️ ورودی باید آرایه باشد');
        return;
    }

    elementIds.forEach(id => {
        initPersianDatepicker(id);
    });
}

/**
 * تبدیل تاریخ شمسی به میلادی (برای ارسال به سرور)
 * @param {string} persianDate - تاریخ شمسی به فرمت YYYY/MM/DD
 * @returns {string} تاریخ میلادی
 */
function persianToGregorian(persianDate) {
    try {
        const parts = persianDate.split('/');
        if (parts.length !== 3) return persianDate;

        const year = parseInt(parts[0]);
        const month = parseInt(parts[1]) - 1;
        const day = parseInt(parts[2]);

        const pd = new persianDate([year, month, day]);
        return pd.toGregorian().format('YYYY-MM-DD');
    } catch (e) {
        console.error('خطا در تبدیل تاریخ:', e);
        return persianDate;
    }
}

/**
 * دریافت تاریخ امروز به شمسی
 * @param {string} format - فرمت خروجی (پیش‌فرض: YYYY/MM/DD)
 * @returns {string}
 */
function getTodayPersian(format = 'YYYY/MM/DD') {
    try {
        return new persianDate().format(format);
    } catch (e) {
        console.error('خطا در دریافت تاریخ امروز:', e);
        return '';
    }
}
/*
// هنگام بارگذاری صفحه
document.addEventListener('DOMContentLoaded', function() {
    // راه‌اندازی یک تقویم
    initPersianDatepicker('createDate');
    
    // راه‌اندازی چند تقویم
    initMultipleDatepickers(['createDate', 'visitDate', 'birthDate']);
    
    // گوش دادن به رویداد انتخاب تاریخ
    document.getElementById('createDate').addEventListener('dateSelected', function(e) {
        console.log('تاریخ انتخاب شد:', e.detail.formattedDate);
    });
});
*/
//==========================================================
//==========================================================
//==========================================================
//==========================================================
