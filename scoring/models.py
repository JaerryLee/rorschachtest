from django.db import models
from accounts.models import User
from django.db.models import Q


class DataTable(models.Model):
    # class Meta:
    #    app_label = 'scoring'
    column1 = models.CharField(max_length=100)
    column2 = models.TextField(max_length=100)
    column3 = models.CharField(max_length=100)
    column4 = models.CharField(max_length=100)
    column5 = models.CharField(max_length=100)
    column6 = models.CharField(max_length=100)
    column7 = models.CharField(max_length=100)
    column8 = models.CharField(max_length=100)
    column9 = models.CharField(max_length=100)
    column10 = models.CharField(max_length=100)
    column11 = models.CharField(max_length=100)
    column12 = models.CharField(max_length=100)
    column13 = models.CharField(max_length=100)
    column14 = models.CharField(max_length=100)
    column15 = models.CharField(max_length=100)

    # def __str__(self):
    #    return f"{self.column1} - {self.column2} - ... - {self.column15}"


class SearchReference(models.Model):
    id = models.CharField(max_length=100, primary_key=True, blank=False)
    Card = models.CharField(max_length=5, blank=False, null=True)
    LOC = models.CharField(max_length=5, null=True, blank=False)
    Cont = models.CharField(max_length=7, null=True, blank=False)
    FQ = models.CharField(max_length=1, null=True, blank=True)
    Determinants = models.CharField(max_length=200, null=True, blank=False)
    Item = models.CharField(max_length=200, null=True, blank=False)
    V = models.CharField(max_length=1, null=True, blank=True)


class CardImages(models.Model):
    card_number = models.CharField(max_length=5, null=True)
    section = models.CharField(max_length=3, null=True)
    img_file = models.FileField(upload_to='images/card', blank=True)
    detail_img = models.FileField(upload_to='images/location', blank=True)


class PopularResponse(models.Model):
    id = models.CharField(max_length=10, primary_key=True, blank=False)
    card_number = models.CharField(max_length=5, null=True)
    p = models.CharField(max_length=200, null=True, blank=False)
    Z = models.CharField(max_length=50, null=True)


class Client(models.Model):
    GENDER_CHOICES = (
        ('M', '남성'),
        ('F', '여성'),
        ('O', '기타'),
    )
    tester = models.ForeignKey(User, on_delete=models.CASCADE, verbose_name='검사자')
    name = models.CharField(max_length=255, verbose_name='이름')
    gender = models.CharField(max_length=1, choices=GENDER_CHOICES, verbose_name='성별')
    birthdate = models.DateField(verbose_name='생년월일', help_text="YYYY-DD-MM 형식으로 입력(예: 2000-08-01)")
    testDate = models.DateField(verbose_name='검사일', help_text="YYYY-DD-MM 형식으로 입력(예: 2023-05-01)")
    notes = models.TextField(blank=True, verbose_name='비고')
    age = models.IntegerField(verbose_name="검사 당시 나이", blank=True, null=True)
    consent = models.BooleanField(default=False)

    # assesor = models.ForeignKey("User", related_name="user", on_delete=models.CASCADE, db_column="username")
    def calculate_age(self):
        test_date = self.testDate
        birth_date = self.birthdate
        if (test_date.month < birth_date.month or
                (test_date.month == birth_date.month and test_date.day < birth_date.day)):
            self.age = test_date.year - birth_date.year - 1
        else:
            self.age = test_date.year - birth_date.year

    def save(self, *args, **kwargs):
        self.calculate_age()
        super().save(*args, **kwargs)


class ResponseCode(models.Model):
    client = models.ForeignKey(Client, on_delete=models.CASCADE, verbose_name='수검자', related_name='responses')
    card = models.CharField(max_length=5, verbose_name='카드번호', null=True)
    response_num = models.IntegerField(verbose_name='반응번호', null=True)
    time = models.CharField(max_length=100, verbose_name='반응시간', blank=True, null=True)
    response = models.TextField(verbose_name='반응내용', blank=True, null=True)
    rotation = models.CharField(max_length=1, verbose_name='카드방향', blank=True, null=True)
    inquiry = models.TextField(verbose_name='질문', blank=True, null=True)
    location = models.CharField(max_length=5, verbose_name='반응영역', blank=True, null=True)
    loc_num = models.IntegerField(verbose_name='영역번호', null=True, blank=True)
    dev_qual = models.CharField(max_length=5, verbose_name='발달질', blank=True, null=True)
    determinants = models.CharField(max_length=30, verbose_name='결정인', blank=True, null=True)
    pair = models.CharField(max_length=20, verbose_name="쌍반응", blank=True, null=True)
    form_qual = models.CharField(max_length=2, verbose_name='형태질', blank=True, null=True)
    content = models.CharField(max_length=30, verbose_name='내용인', blank=True, null=True)
    popular = models.CharField(max_length=1, verbose_name='평범반응', blank=True, null=True)
    Z = models.CharField(max_length=5, verbose_name='조직화 점수', blank=True, null=True)
    special = models.CharField(max_length=50, verbose_name='특수점수', blank=True, null=True)
    comment = models.TextField(verbose_name='코멘트', blank=True, null=True)

    def __str__(self):
        return self.client.name


class StructuralSummary(models.Model):
    client = models.ForeignKey('Client', on_delete=models.CASCADE, verbose_name='수검자')

    # 1. Location Features
    # 1-1. 조직화 활동
    Zf = models.PositiveIntegerField(verbose_name='Zf', default=0)
    Zsum = models.PositiveIntegerField(verbose_name='Zsum', default=0)
    Zest = models.FloatField(verbose_name='Zest', default=0.0)  # 0이면 NA로 표시
    # 1-2. 반응영역의 빈도
    W = models.PositiveIntegerField(verbose_name='W', default=0)
    D = models.PositiveIntegerField(verbose_name='D', default=0)
    Dd = models.PositiveIntegerField(verbose_name='Dd', default=0)
    S = models.PositiveIntegerField(verbose_name='S', default=0)

    # 2. Developmental Quality
    dev_plus = models.PositiveIntegerField(verbose_name='dev_+', default=0)
    dev_o = models.PositiveIntegerField(verbose_name='dev_o', default=0)
    dev_vplus = models.PositiveIntegerField(verbose_name='dev_v/+', default=0)
    dev_v = models.PositiveIntegerField(verbose_name='dev_v', default=0)

    # 3. Form Quality
    # 3-1. FQx
    fqx_plus = models.PositiveIntegerField(verbose_name='fqx_+', default=0)
    fqx_o = models.PositiveIntegerField(verbose_name='fqx_o', default=0)
    fqx_u = models.PositiveIntegerField(verbose_name='fqx_u', default=0)
    fqx_minus = models.PositiveIntegerField(verbose_name='fqx_minus', default=0)
    fqx_none = models.PositiveIntegerField(verbose_name='fqx_none', default=0)
    # 3-2. Mqaul
    mq_plus = models.PositiveIntegerField(verbose_name='mq_+', default=0)
    mq_o = models.PositiveIntegerField(verbose_name='mq_o', default=0)
    mq_u = models.PositiveIntegerField(verbose_name='mq_u', default=0)
    mq_minus = models.PositiveIntegerField(verbose_name='mq_minus', default=0)
    mq_none = models.PositiveIntegerField(verbose_name='mq_none', default=0)
    # 3-3. W+D
    wd_plus = models.PositiveIntegerField(verbose_name='wd_+', default=0)
    wd_o = models.PositiveIntegerField(verbose_name='wd_o', default=0)
    wd_u = models.PositiveIntegerField(verbose_name='wd_u', default=0)
    wd_minus = models.PositiveIntegerField(verbose_name='wd_minus', default=0)
    wd_none = models.PositiveIntegerField(verbose_name='wd_none', default=0)

    # 4. Determinants
    # 4-1. 혼합 결정인
    blends = models.CharField(max_length=100, default='')
    # 4-2. 단일 변인
    M = models.PositiveIntegerField(db_column='M', verbose_name='M', null=True)
    FM = models.PositiveIntegerField(verbose_name='FM', null=True)
    m_l = models.PositiveIntegerField(db_column='m_l', verbose_name="m'", null=True)
    FC = models.PositiveIntegerField(verbose_name='FC', null=True)
    CF = models.PositiveIntegerField(verbose_name='CF', null=True)
    C = models.PositiveIntegerField(verbose_name='C', null=True, blank=True)
    Cn = models.PositiveIntegerField(verbose_name='Cn', null=True)
    FCa = models.PositiveIntegerField(verbose_name="FC'", null=True)
    CaF = models.PositiveIntegerField(verbose_name="C'F", null=True)
    Ca = models.PositiveIntegerField(verbose_name="C'", null=True)
    FT = models.PositiveIntegerField(verbose_name='FT', null=True)
    TF = models.PositiveIntegerField(verbose_name='TF', null=True)
    T = models.PositiveIntegerField(verbose_name='T', null=True)
    FV = models.PositiveIntegerField(verbose_name='FV', null=True)
    VF = models.PositiveIntegerField(verbose_name='VF', null=True)
    V = models.PositiveIntegerField(verbose_name='V', null=True)
    FY = models.PositiveIntegerField(verbose_name='FY', null=True)
    YF = models.PositiveIntegerField(verbose_name='YF', null=True)
    Y = models.PositiveIntegerField(verbose_name='Y', null=True)
    Fr = models.PositiveIntegerField(verbose_name='Fr', null=True)
    rF = models.PositiveIntegerField(verbose_name='rF', null=True)
    FD = models.PositiveIntegerField(verbose_name='FD', null=True)
    F = models.PositiveIntegerField(verbose_name='F', null=True)
    pair = models.PositiveIntegerField(verbose_name='(2)', null=True)

    # 5. 내용인
    H = models.PositiveIntegerField(verbose_name='H', default=0)
    H_paren = models.PositiveIntegerField(verbose_name='(H)', default=0)
    Hd = models.PositiveIntegerField(verbose_name='Hd', default=0)
    Hd_paren = models.PositiveIntegerField(verbose_name='(Hd)', default=0)
    Hx = models.PositiveIntegerField(verbose_name='Hx', default=0)
    A = models.PositiveIntegerField(verbose_name='A', default=0)
    A_paren = models.PositiveIntegerField(verbose_name='(A)', default=0)
    Ad = models.PositiveIntegerField(verbose_name='Ad', default=0)
    Ad_paren = models.PositiveIntegerField(verbose_name='(Ad)', default=0)
    An = models.PositiveIntegerField(verbose_name='An', default=0)
    Art = models.PositiveIntegerField(verbose_name='Art', default=0)
    Ay = models.PositiveIntegerField(verbose_name='Ay', default=0)
    Bl = models.PositiveIntegerField(verbose_name='Bl', default=0)
    Bt = models.PositiveIntegerField(verbose_name='Bt', default=0)
    Cg = models.PositiveIntegerField(verbose_name='Cg', default=0)
    Cl = models.PositiveIntegerField(verbose_name='Cl', default=0)
    Ex = models.PositiveIntegerField(verbose_name='Ex', default=0)
    Fd_l = models.PositiveIntegerField(verbose_name='Fd_l', default=0)
    Fi = models.PositiveIntegerField(verbose_name='Fi', default=0)
    Ge = models.PositiveIntegerField(verbose_name='Ge', default=0)
    Hh = models.PositiveIntegerField(verbose_name='Hh', default=0)
    Ls = models.PositiveIntegerField(verbose_name='Ls', default=0)
    Na = models.PositiveIntegerField(verbose_name='Na', default=0)
    Sc = models.PositiveIntegerField(verbose_name='Sc', default=0)
    Sx = models.PositiveIntegerField(verbose_name='Sx', default=0)
    Xy = models.PositiveIntegerField(verbose_name='Xy', default=0)
    Idio = models.PositiveIntegerField(verbose_name='Id', default=0)

    # 6. approach 인지적 접근 방식
    app_I = models.CharField(max_length=100, default='')
    app_II = models.CharField(max_length=100, default='')
    app_III = models.CharField(max_length=100, default='')
    app_IV = models.CharField(max_length=100, default='')
    app_V = models.CharField(max_length=100, default='')
    app_VI = models.CharField(max_length=100, default='')
    app_VII = models.CharField(max_length=100, default='')
    app_VIII = models.CharField(max_length=100, default='')
    app_IX = models.CharField(max_length=100, default='')
    app_X = models.CharField(max_length=100, default='')

    # 7. 특수 점수
    sp_dv = models.PositiveIntegerField(verbose_name='sp_DV', default=0)
    sp_dv2 = models.PositiveIntegerField(verbose_name='sp_DV2', default=0)
    sp_dr = models.PositiveIntegerField(verbose_name='sp_DR', default=0)
    sp_dr2 = models.PositiveIntegerField(verbose_name='sp_DR2', default=0)
    sp_inc = models.PositiveIntegerField(verbose_name='sp_INC', default=0)
    sp_inc2 = models.PositiveIntegerField(verbose_name='sp_INC2', default=0)
    sp_fab = models.PositiveIntegerField(verbose_name='sp_FAB', default=0)
    sp_fab2 = models.PositiveIntegerField(verbose_name='sp_FAB2', default=0)
    sp_alog = models.PositiveIntegerField(verbose_name='sp_ALOG', default=0)
    sp_con = models.PositiveIntegerField(verbose_name='sp_CON', default=0)
    sum6 = models.PositiveIntegerField(verbose_name='raw_sum6', default=0)
    wsum6 = models.PositiveIntegerField(verbose_name='wgtd_sum6', default=0)
    sp_psv = models.PositiveIntegerField(verbose_name='sp_PSV', default=0)
    sp_ab = models.PositiveIntegerField(verbose_name='sp_AB', default=0)
    sp_ag = models.PositiveIntegerField(verbose_name='sp_AG', default=0)
    sp_cop = models.PositiveIntegerField(verbose_name='sp_COP', default=0)
    sp_mor = models.PositiveIntegerField(verbose_name='sp_MOR', default=0)
    sp_per = models.PositiveIntegerField(verbose_name='sp_PER', default=0)
    sp_cp = models.PositiveIntegerField(verbose_name='sp_CP', default=0)
    sp_ghr = models.PositiveIntegerField(verbose_name='sp_GHR', default=0)
    sp_phr = models.PositiveIntegerField(verbose_name='sp_PHR', default=0)

    # 8. CORE 핵심 영역
    # 8-1. R, L
    R = models.IntegerField(verbose_name='반응수')
    L = models.FloatField(verbose_name='lambda', default=0)
    # 8-2. 스트레스 및 통제 관련 영역
    ErleBnistypus = models.CharField(max_length=20, verbose_name='Erlebnistypus', default="")
    EA = models.FloatField(verbose_name='EA', default=0)
    EBper = models.FloatField(verbose_name='EBper', default=0)  # 0이면 NA로 표시
    eb = models.CharField(max_length=20, verbose_name='eb', default="")
    es = models.FloatField(verbose_name='es', default=0)
    D_score = models.IntegerField(verbose_name='D', default=0)
    adj_es = models.FloatField(verbose_name='es', default=0)
    adj_D = models.IntegerField(verbose_name='D', default=0)
    # 8-3. 결정인들
    sum_FM = models.PositiveIntegerField(verbose_name='sumFM', default=0)
    sum_m = models.PositiveIntegerField(verbose_name='sum_m', default=0)
    sum_Ca = models.PositiveIntegerField(verbose_name="sum_C'", default=0)
    sum_V = models.PositiveIntegerField(verbose_name="sum_V", default=0)
    sum_T = models.PositiveIntegerField(verbose_name="sum_T", default=0)
    sum_Y = models.PositiveIntegerField(verbose_name="sum_Y", default=0)

    # 9. 정서 영역
    f_c_prop = models.CharField(max_length=20, verbose_name="FC:CF+C", default='')
    pure_c = models.IntegerField(verbose_name="pure C", default=0)
    ca_c_prop = models.CharField(max_length=20, verbose_name="SumC':WsumC", default='')
    afr = models.FloatField(verbose_name='affection ratio', default=0.0)
    # 공백반응은 location features의 S 사용
    blends_r = models.CharField(max_length=20, verbose_name="Blends:R", default='')
    # 색채투사는 특수점수의 sp_cp 사용

    # 10. 대인관계
    # cop, ag, per은 특수점수의 sp_cop, sp_ag, sp_per 사용
    # food는 내용인의 fd_l 사용, pure H는 내용인의 H 사용
    # sum_T는 핵심영역의 sum_T 사용
    GHR_PHR = models.CharField(max_length=20, verbose_name='GHR:PHR', default='')
    a_p = models.CharField(max_length=20, verbose_name='a:p', default='')
    human_cont = models.PositiveIntegerField(verbose_name='human_content', default=0)
    Isol = models.FloatField(verbose_name='Isol Index', default=0.0)

    # 11. Ideation
    # a:p는 대인관계의 a_p 이용, sum6, wsum6은 특수점수의 sum6, wsum6 이용, MOR도 특수점수의 sp_mor 이용
    # M-와 M-none은 mqual_minus, mqual_none 사용
    Ma_Mp = models.CharField(max_length=20, verbose_name="Ma:Mp", default='')
    Lvl_2 = models.PositiveIntegerField(verbose_name="lvl-2", default=0)
    intel = models.FloatField(verbose_name='intellecutalization index', default=0.0)

    # 12. Mediation
    x_minus_per = models.FloatField(verbose_name="X-%", default=0.0)
    xa_per = models.FloatField(verbose_name="XA%", default=0.0)
    wda_per = models.FloatField(verbose_name="WDA%", default=0.0)
    s_minus = models.PositiveIntegerField(verbose_name="S-%", default=0)
    popular = models.PositiveIntegerField(verbose_name='P', default=0)
    x_plus_per = models.FloatField(verbose_name='X+%', default=0.0)
    xu_per = models.FloatField(verbose_name='Xu%', default=0.0)

    # 13. Processing
    # Zf, sp_psv, dev_plus, dev_v
    Zd = models.FloatField(verbose_name='Zd', default=0.0)  # 0이면 NA로 표시
    W_D_Dd = models.CharField(max_length=15, verbose_name="W:D:Dd", default='')
    W_M = models.CharField(max_length=15, verbose_name="W:M", default='')

    # 14. self perception
    # SumV는 Sum_V 사용, Mor는 sp_mor 사용
    ego = models.FloatField(verbose_name="egocentric", default=0.0)
    fr_rf = models.PositiveIntegerField(verbose_name="Fr+rF", default=0)
    fdn = models.PositiveIntegerField(verbose_name="selfperception_FD", default=0)
    an_xy = models.PositiveIntegerField(verbose_name="An+Xy", default=0)
    h_prop = models.CharField(max_length=15, verbose_name="H:(H)+Hd+(Hd)", default='')

    # 15. 특수 지표
    PTI = models.CharField(max_length=15, verbose_name="PTI", default='xxxxx')
    sumPTI = models.IntegerField(default=0)
    DEPI = models.CharField(max_length=15, verbose_name="DEPI", default='xxxxxxx')
    sumDEPI = models.IntegerField(default=0)
    CDI = models.CharField(max_length=15, verbose_name="CDI", default='xxxxx')
    sumCDI = models.IntegerField(default=0)
    SCON = models.CharField(max_length=15, verbose_name="S-CON", default='xxxxxxxxxx')
    sumSCON = models.IntegerField(default=0)
    HVI_premise = models.BooleanField(default=False)
    HVI = models.CharField(max_length=15, verbose_name="HVI", default='xxxxxxx')
    sumHVI = models.IntegerField(default=0)
    HVI_except = models.CharField(max_length=15, verbose_name="HVI", default='')
    OBS = models.CharField(max_length=15, verbose_name="HVI", default='xxxxxxxxx')
    OBS_posi = models.BooleanField(default=False)

    def calculate_values(self):
        age = self.client.age
        response_codes = ResponseCode.objects.filter(client=self.client)
        determinants_list = []
        all_d_list = []
        contents_list = []
        blends = ''
        special_list = []
        elements_to_check = ['fy', 'yf', 'y', 'ft', 'tf', 't', 'fv', 'vf', 'v', "c'f", "fc'", "c'"]  # shd elements
        col_shd_blends = 0
        shd_blends = 0

        # 1. Location Features
        zf = response_codes.filter(Q(Z='ZD') | Q(Z='ZW') | Q(Z='ZS') | Q(Z='ZA')).count()
        self.Zf = zf

        # 로마숫자 아라비안 숫자로 변환
        roman_dict = {
            'I': '1', 'II': '2', 'III': '3', 'IV': '4', 'V': '5',
            'VI': '6', 'VII': '7', 'VIII': '8', 'IX': '9', 'X': '10'
        }
        for response_code in response_codes:
            response_code.card = roman_dict.get(response_code.card, response_code.card)
            response_code.save()

        # Zsum 계산
        z_sum_dict = {
            '1': {'ZW': 1, 'ZA': 4, 'ZD': 6, 'ZS': 3.5},
            '2': {'ZW': 4.5, 'ZA': 3, 'ZD': 5.5, 'ZS': 4.5},
            '3': {'ZW': 5.5, 'ZA': 3, 'ZD': 4, 'ZS': 4.5},
            '4': {'ZW': 2, 'ZA': 4, 'ZD': 3.5, 'ZS': 5},
            '5': {'ZW': 1, 'ZA': 2.5, 'ZD': 5, 'ZS': 4},
            '6': {'ZW': 2.5, 'ZA': 2.5, 'ZD': 6, 'ZS': 6.5},
            '7': {'ZW': 2.5, 'ZA': 1, 'ZD': 3, 'ZS': 4},
            '8': {'ZW': 4.5, 'ZA': 3, 'ZD': 3, 'ZS': 4},
            '9': {'ZW': 5.5, 'ZA': 2.5, 'ZD': 4.5, 'ZS': 5},
            '10': {'ZW': 5.5, 'ZA': 4, 'ZD': 4.5, 'ZS': 6}
        }
        self.Zsum = sum(
            z_sum_dict.get(response_code.card, {}).get(response_code.Z, 0) for response_code in
            response_codes)
        # Zest 계산
        if zf == 0:
            self.Zest = 0
        elif zf > 50:
            self.Zest = 173
        else:
            # Zf 값이 1부터 50 사이에 있는 경우
            z_est_dict = {
                1: 0, 2: 2.5, 3: 6, 4: 10, 5: 13.5, 6: 17, 7: 20.5, 8: 24, 9: 27.5,
                10: 31, 11: 34.5, 12: 38, 13: 41.5, 14: 45.5, 15: 49, 16: 51.5, 17: 56,
                18: 59.5, 19: 63, 20: 66.5, 21: 70, 22: 73.5, 23: 77, 24: 81, 25: 84.5,
                26: 88, 27: 91.5, 28: 95, 29: 98.5, 30: 102.5, 31: 105.5, 32: 109.5,
                33: 112.5, 34: 116.5, 35: 120, 36: 123.5, 37: 127, 38: 130.5, 39: 134,
                40: 137.5, 41: 141, 42: 144, 43: 148, 44: 152, 45: 155.5, 46: 159,
                47: 162.5, 48: 166, 49: 169.5, 50: 173
            }
            self.Zest = z_est_dict.get(zf, 0)
        if self.Zest > 0:
            self.Zd = self.Zsum - self.Zest
        else:
            self.Zd = 0

        # 'W', 'D', 'Dd', 'S' 빈도 계산
        self.W = response_codes.filter(location__contains='W').count()
        self.D = response_codes.filter(Q(location__exact='D') | Q(location__exact="DS")).count()
        self.Dd = response_codes.filter(location__contains='Dd').count()
        self.S = response_codes.filter(location__contains='S').count()

        # 2. Developmental Quality
        from collections import Counter
        # 발달질 빈도 계산
        dev_qual_list = [response_code.dev_qual for response_code in response_codes if response_code.dev_qual]
        dev_qual_counts = Counter(dev_qual_list)
        self.dev_plus = dev_qual_counts['+']
        self.dev_o = dev_qual_counts['o']
        self.dev_vplus = dev_qual_counts['v/+']
        self.dev_v = dev_qual_counts['v']

        # 3. Form Quality
        # 3-1. FQx
        form_qual_list = [response_code.form_qual for response_code in response_codes if response_code.form_qual]
        form_qual_counts = Counter(form_qual_list)
        self.fqx_plus = form_qual_counts['+']
        self.fqx_o = form_qual_counts['o']
        self.fqx_u = form_qual_counts['u']
        self.fqx_minus = form_qual_counts['-']
        self.fqx_none = form_qual_counts['none']
        # 3-2. Mqual
        mqual_list = response_codes.filter(determinants__regex=r'(?<!F)M')
        mqual_counts = Counter([mqual.form_qual for mqual in mqual_list if mqual.form_qual])
        self.mq_plus = mqual_counts['+']
        self.mq_o = mqual_counts['o']
        self.mq_u = mqual_counts['u']
        self.mq_minus = mqual_counts['-']
        self.mq_none = mqual_counts['none']
        # 3-3. W+D
        wd_list = response_codes.filter(location__regex=r'\b(W|D|DS|WS)\b')
        wd_counts = Counter(wd.form_qual for wd in wd_list if wd.form_qual)
        self.wd_plus = wd_counts['+']
        self.wd_o = wd_counts['o']
        self.wd_u = wd_counts['u']
        self.wd_minus = wd_counts['-']
        self.wd_none = wd_counts['no']

        # Collect all determinants from ResponseCode objects
        import re
        for response_code in response_codes:
            # 결정인
            determinants = re.split(r'[.,]+', response_code.determinants.replace(' ', ''))
            determinants_lower = [item.lower() if item not in ['ma', 'mp', 'Ma', 'Mp', 'Ma-p']
                                  else item for item in determinants]
            all_d_list.extend(determinants_lower)
            # 혼합반응
            if len(determinants) >= 2:
                blends += '.'.join(determinants) + ','
                if (any(element in determinants_lower for element in elements_to_check)
                        and any(element in determinants_lower for element in ['c', 'cf', 'fc'])):
                    col_shd_blends += 1
                if len(set(elements_to_check) & set(determinants_lower)) >= 2:
                    shd_blends += 1
            else:
                determinants_list.extend(determinants_lower)  # single list
            # 쌍반응
            if response_code.pair:
                determinants_list.append(response_code.pair)
            # 내용
            contents = re.split(r'[.,]+', response_code.content.replace(' ', ''))
            contents_lower = [item.lower() for item in contents]
            contents_list.extend(contents_lower)
            # 7. 특수점수
            if response_code.special is not None:
                specials = re.split(r'[.,]+', response_code.special.replace(' ', ''))
                specials = ' '.join(specials).split()
                specials = [value for value in specials if value not in ["GHR", "PHR"]]
            else:
                specials = []
            if (any(content in contents_lower for content in ['h', '(h)', 'hd', '(hd)', 'hx'])
                    or any(determinant in determinants_lower for determinant in ['Ma', 'Mp', 'Ma-p'])
                    or (any(determinant in determinants_lower for determinant in ['fma', 'fmp', 'fma-p'])
                        and any(special in specials for special in ['COP', 'AG']))):
                if ("h" in contents_lower and response_code.form_qual in ["+", "o", "u"]
                        and all(special not in specials for special in
                                ['DV2', 'DR', 'DR2', 'INC', 'INC2', 'FAB', 'FAB2', 'CON', 'ALOG', 'AG', 'MOR'])):
                    specials.append("GHR")
                elif response_code.form_qual in ["-", "no"] or any(
                        special in specials for special in ['DV2', 'DR2', 'INC2', 'FAB2', 'CON', 'ALOG']):
                    specials.append("PHR")
                elif "COP" in specials and "AG" not in specials:
                    specials.append("GHR")
                elif "FAB" in specials or "MOR" in specials or "an" in contents_lower:
                    specials.append("PHR")
                elif response_code.popular == "P" and response_code.card in ['3', '4', '7', '9']:
                    specials.append("GHR")
                elif any(special in specials for special in ['AG', 'INC', 'DR']) or "hd" in contents_lower:
                    specials.append("PHR")
                else:
                    specials.append("GHR")
            special_list.extend(specials)
            response_code.special = ','.join(specials)
            response_code.save()

        # 4-1. 혼합 결정인
        self.blends = blends
        # 4-2. 단일 결정인 빈도 계산
        self.M = determinants_list.count('Ma') + determinants_list.count('Mp') + determinants_list.count('Ma-p')
        self.FM = determinants_list.count('fma') + determinants_list.count('fmp') + determinants_list.count('fma-p')
        self.m_l = determinants_list.count('ma') + determinants_list.count('mp') + determinants_list.count('ma-p')
        self.FC = determinants_list.count('fc')
        self.CF = determinants_list.count('cf')
        self.C = determinants_list.count('c')
        self.Cn = determinants_list.count('cn')
        self.FCa = determinants_list.count("fc'")
        self.CaF = determinants_list.count("c'f")
        self.Ca = determinants_list.count("c'")
        self.FT = determinants_list.count('ft')
        self.TF = determinants_list.count('tf')
        self.T = determinants_list.count('t')
        self.FV = determinants_list.count('fv')
        self.VF = determinants_list.count('vf')
        self.V = determinants_list.count('v')
        self.FY = determinants_list.count('fy')
        self.YF = determinants_list.count('yf')
        self.Y = determinants_list.count('y')
        self.Fr = determinants_list.count('fr')
        self.rF = determinants_list.count('rf')
        self.FD = determinants_list.count('fd')
        self.F = determinants_list.count('f')
        self.pair = determinants_list.count('2')

        # 5. 내용인 빈도 계산
        self.H = contents_list.count('h')
        self.H_paren = contents_list.count('(h)')
        self.Hd = contents_list.count('hd')
        self.Hd_paren = contents_list.count('(hd)')
        self.Hx = contents_list.count('hx')
        self.A = contents_list.count('a')
        self.A_paren = contents_list.count('(a)')
        self.Ad = contents_list.count('ad')
        self.Ad_paren = contents_list.count('(ad)')
        self.An = contents_list.count('an')
        self.Art = contents_list.count('art')
        self.Ay = contents_list.count('ay')
        self.Bl = contents_list.count('bl')
        self.Bt = contents_list.count('bt')
        self.Cg = contents_list.count('cg')
        self.Cl = contents_list.count('cl')
        self.Ex = contents_list.count('ex')
        self.Fd_l = contents_list.count('fd')
        self.Fi = contents_list.count('fi')
        self.Ge = contents_list.count('ge')
        self.Hh = contents_list.count('hh')
        self.Ls = contents_list.count('ls')
        self.Na = contents_list.count('na')
        self.Sc = contents_list.count('sc')
        self.Sx = contents_list.count('sx')
        self.Xy = contents_list.count('xy')
        self.Idio = contents_list.count('id')

        # 6. approach 인지적 접근 방식
        app_data = {'1': [], '2': [], '3': [], '4': [], '5': [], '6': [], '7': [], '8': [], '9': [], '10': []}
        for response_code in response_codes:
            app_data[response_code.card].append(response_code.location)
        for arab, rom in zip(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'],
                             ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']):
            setattr(self, f'app_{rom}', '.'.join(app_data[arab]))

        # 7. 특수점수
        special_counts = Counter(special_list)
        self.sp_dv = special_counts['DV']
        self.sp_dv2 = special_counts['DV2']
        self.sp_dr = special_counts['DR']
        self.sp_dr2 = special_counts['DR2']
        self.sp_inc = special_counts['INC']
        self.sp_inc2 = special_counts['INC2']
        self.sp_fab = special_counts['FAB']
        self.sp_fab2 = special_counts['FAB2']
        self.sp_alog = special_counts['ALOG']
        self.sp_con = special_counts['CON']
        self.sp_psv = special_counts['PSV']
        self.sp_ab = special_counts['AB']
        self.sp_ag = special_counts['AG']
        self.sp_cop = special_counts['COP']
        self.sp_mor = special_counts['MOR']
        self.sp_per = special_counts['PER']
        self.sp_cp = special_counts['CP']
        self.sp_ghr = special_counts['GHR']
        self.sp_phr = special_counts['PHR']
        self.sum6 = (
                self.sp_dv + self.sp_dv2 + self.sp_dr + self.sp_dr2 +
                self.sp_inc + self.sp_inc2 + self.sp_fab + self.sp_fab2 +
                self.sp_alog + self.sp_con
        )
        self.wsum6 = (
                (1 * self.sp_dv) + (2 * self.sp_dv2) + (2 * self.sp_inc) +
                (4 * self.sp_inc2) + (3 * self.sp_dr) + (6 * self.sp_dr2) +
                (4 * self.sp_fab) + (7 * self.sp_fab2) + (5 * self.sp_alog) +
                (7 * self.sp_con)
        )

        # 8. 핵심 영역
        # 8-1. R, L
        self.R = len(response_codes)
        if (self.R - self.F) != 0:
            self.L = self.F / (self.R - self.F)
        else:
            self.L = self.F / (self.R - self.F + 0.001)
        # 8-2. 통제 변인 8-3. 결정인들
        sum_M = all_d_list.count('Ma') + all_d_list.count('Mp') + all_d_list.count('Ma-p')
        sum_fc = all_d_list.count('fc')
        sum_cf = all_d_list.count('cf')
        sum_c = all_d_list.count('c')
        wsumc = 0.5 * sum_fc + 1 * sum_cf + 1.5 * sum_c
        self.ErleBnistypus = "{}:{}".format(sum_M, wsumc)
        self.EA = sum_M + wsumc
        self.sum_FM = all_d_list.count('fma') + all_d_list.count('fmp') + all_d_list.count('fma-p')
        self.sum_m = all_d_list.count('ma') + all_d_list.count('mp') + all_d_list.count('ma-p')
        self.sum_Ca = all_d_list.count("c'") + all_d_list.count("fc'") + all_d_list.count("c'f")
        self.sum_V = all_d_list.count('v') + all_d_list.count('vf') + all_d_list.count('fv')
        self.sum_T = all_d_list.count('t') + all_d_list.count('tf') + all_d_list.count('ft')
        self.sum_Y = all_d_list.count('y') + all_d_list.count('yf') + all_d_list.count('fy')
        sumshading = self.sum_Ca + self.sum_T + self.sum_V + self.sum_Y
        sumFm_m = self.sum_FM + self.sum_m
        self.eb = "{}:{}".format(sumFm_m, sumshading)
        self.es = sumFm_m + sumshading
        extra_m = (self.sum_m - 1) if self.sum_m > 1 else 0
        extra_Y = (self.sum_Y - 1) if self.sum_Y > 1 else 0
        self.adj_es = self.es - (extra_m + extra_Y)
        self.D_score = int((self.EA - self.es) / 2.5000001)
        self.adj_D = int((self.EA - self.adj_es) / 2.5000001)
        if self.EA >= 4.0 and self.L < 1.0:
            if (4.0 <= self.EA <= 10.0 and abs(sum_M - wsumc) > 2.0) or (self.EA > 10.0 and abs(sum_M - wsumc) > 2.5):
                self.EBper = max(float(sum_M), wsumc) / min(float(sum_M), wsumc)
            else:
                self.EBper = 0
        else:
            self.EBper = 0

        # 9. 정서 영역
        self.f_c_prop = "{}:{}".format(sum_fc, sum_cf + sum_c)
        self.pure_c = sum_c
        self.ca_c_prop = "{}:{}".format(self.sum_Ca, wsumc)
        afr_count_numerator = response_codes.filter(Q(card='8') | Q(card='9') | Q(card='10')).count()
        afr_count_denominator = response_codes.filter(
            Q(card='1') | Q(card='2') | Q(card='3') | Q(card='4') | Q(card='5') | Q(card='6') | Q(card='7')).count()
        # Calculate affection ratio (afr)
        if afr_count_denominator != 0:
            self.afr = afr_count_numerator / afr_count_denominator
        else:
            self.afr = 1  # 불가능한 값
        blends = blends.rstrip(',')  # blends_num 계산 시 맨 오른쪽 쉼표 제거
        blends_num = len(blends.split(','))
        self.blends_r = "{}:{}".format(blends_num, self.R)

        # 10. 대인관계
        self.GHR_PHR = "{}:{}".format(self.sp_ghr, self.sp_phr)
        sum_a = (all_d_list.count('ma') + all_d_list.count('Ma') + all_d_list.count('fma') +
                 all_d_list.count('ma-p') + all_d_list.count('Ma-p') + all_d_list.count('fma-p'))
        sum_p = (all_d_list.count('mp') + all_d_list.count('Mp') + all_d_list.count('fmp') +
                 all_d_list.count('ma-p') + all_d_list.count('Ma-p') + all_d_list.count('fma-p'))
        self.a_p = "{}:{}".format(sum_a, sum_p)
        self.human_cont = self.H + self.H_paren + self.Hd + self.Hd_paren
        self.Isol = (self.Bt + 2 * self.Cl + self.Ge + self.Ls + 2 * self.Na) / self.R

        # 11. Ideation
        sum_Ma = all_d_list.count('Ma') + all_d_list.count('Ma-p')
        sum_Mp = all_d_list.count('Mp') + all_d_list.count('Ma-p')
        self.Ma_Mp = "{}:{}".format(sum_Ma, sum_Mp)
        self.Lvl_2 = self.sp_dv2 + self.sp_dr2 + self.sp_inc2 + self.sp_fab2
        self.intel = 2 * self.sp_ab + self.Art + self.Ay

        # 12. Mediation
        self.x_minus_per = self.fqx_minus / self.R
        self.xa_per = (self.fqx_plus + self.fqx_o + self.fqx_u) / self.R
        self.wda_per = (self.wd_plus + self.wd_o + self.wd_u) / (
                self.wd_plus + self.wd_o + self.wd_u + self.wd_minus + self.wd_none)
        s_list = response_codes.filter(location__contains='S')
        s_counts = Counter(s.form_qual for s in s_list if s.form_qual)
        self.s_minus = s_counts['-']
        self.popular = response_codes.filter(popular__contains="P").count()
        self.x_plus_per = (self.fqx_plus + self.fqx_o) / self.R
        self.xu_per = self.fqx_u / self.R

        # 13. Processing
        self.W_D_Dd = "{}:{}:{}".format(self.W, self.D, self.Dd)
        self.W_M = "{}:{}".format(self.W, sum_M)

        # 14. Self perception
        r = all_d_list.count('fr') + all_d_list.count('rf')
        self.ego = (3 * r + self.pair) / self.R
        self.fr_rf = r
        self.fdn = all_d_list.count('fd')
        self.an_xy = self.An + self.Xy
        self.h_prop = "{}:{}".format(self.H, self.H_paren + self.Hd + self.Hd_paren)

        # 15. 특수지표
        # 아동청소년의 경우 조정된 자기중심성 지표와 WSum6, Afr criteria
        highR_wsum6_map = {
            (5, 7): 20,
            (8, 10): 19,
            (11, 13): 18,
            (14, float('inf')): 16  # 14세 이상은 17으로 처리
        }
        for age_range, value in highR_wsum6_map.items():
            if age_range[0] <= age <= age_range[1]:
                highR_wsum6_crt = value
                break
        else:
            highR_wsum6_crt = 20  # 5세 미만은 20으로 일단 처리
        lowR_wsum6_map = {
            (5, 7): 16,
            (8, 10): 15,
            (11, 13): 14,
            (14, float('inf')): 12  # 14세 이상은 12로 처리
        }
        for age_range, value in lowR_wsum6_map.items():
            if age_range[0] <= age <= age_range[1]:
                lowR_wsum6_crt = value
                break
        else:
            lowR_wsum6_crt = 16  # 5세 미만은 16으로 일단 처리
        afr_map = {
            (5, 6): 0.57,
            (7, 9): 0.55,
            (10, 13): 0.53,
            (14, float('inf')): 0.46
        }
        for age_range, value in afr_map.items():
            if age_range[0] <= age <= age_range[1]:
                afr_crt = value
                break
        else:
            afr_crt = 0.57  # 5세 미만은 0.57으로 일단 처리
        ego_map = {
            5: (0.55, 0.83),
            6: (0.52, 0.82),
            7: (0.52, 0.77),
            8: (0.48, 0.74),
            9: (0.45, 0.69),
            10: (0.45, 0.63),
            11: (0.45, 0.58),
            12: (0.38, 0.58),
            13: (0.38, 0.56),
            14: (0.37, 0.54),
            15: (0.33, 0.5),
            16: (0.33, 0.48)
        }
        if age in ego_map:
            high_ego_crt, low_ego_crt = ego_map[age]
        elif age >= 17:
            high_ego_crt, low_ego_crt = 0.33, 0.44
        else:
            high_ego_crt, low_ego_crt = 0.55, 0.83
        PTI1 = "o" if self.xa_per < 0.70 and self.wda_per < 0.75 else "x"
        PTI2 = "o" if self.x_minus_per > 0.29 else "x"
        PTI3 = "o" if self.Lvl_2 > 2 and self.sp_fab2 > 0 else "x"
        PTI4 = "o" if (self.R < 17 and self.wsum6 > lowR_wsum6_crt) or (
                self.R > 16 and self.wsum6 > highR_wsum6_crt) else "x"
        PTI5 = "o" if self.mq_minus > 1 or self.x_minus_per > 0.40 else "x"
        self.PTI = PTI1 + PTI2 + PTI3 + PTI4 + PTI5
        self.sumPTI = self.PTI.count("o")
        DEPI1 = "o" if self.sum_V > 0 or self.fdn > 2 else "x"
        DEPI2 = "o" if col_shd_blends > 0 or self.S > 2 else "x"
        DEPI3 = "o" if (self.ego > high_ego_crt and self.fr_rf == 0) or (self.ego < low_ego_crt) else "x"
        DEPI4 = "o" if self.afr < afr_crt or blends_num < 4 else "x"
        DEPI5 = "o" if sumshading > sumFm_m or self.sum_Ca > 2 else "x"
        DEPI6 = "o" if self.sp_mor > 2 or self.intel > 3 else "x"
        DEPI7 = "o" if self.sp_cop < 2 or self.Isol > 0.24 else "x"
        self.DEPI = DEPI1 + DEPI2 + DEPI3 + DEPI4 + DEPI5 + DEPI6 + DEPI7
        self.sumDEPI = self.DEPI.count("o")  # 5이상
        CDI1 = "o" if self.EA < 6 or self.adj_D < 0 else "x"
        CDI2 = "o" if self.sp_cop < 2 and self.sp_ag < 2 else "x"
        CDI3 = "o" if wsumc < 2.5 or self.afr < afr_crt else "x"
        CDI4 = "o" if sum_p > sum_a + 1 or self.H < 2 else "x"
        CDI5 = "o" if self.sum_T > 1 or self.Isol > 0.24 or self.Fd_l > 0 else "x"
        self.CDI = CDI1 + CDI2 + CDI3 + CDI4 + CDI5
        self.sumCDI = self.CDI.count("o")  # 4이상
        SCON1 = "o" if self.sum_V + self.fdn > 2 else "x"
        SCON2 = "o" if col_shd_blends > 0 else "x"
        SCON3 = "o" if self.ego < 0.31 or self.ego > 0.44 else "x"
        SCON4 = "o" if self.sp_mor > 3 else "x"
        SCON5 = "o" if self.Zd > 3.5 or self.Zd < -3.5 else "x"
        SCON6 = "o" if self.es > self.EA else "x"
        SCON7 = "o" if sum_cf + sum_c > sum_fc else "x"
        SCON8 = "o" if self.x_plus_per < 0.70 else "x"
        SCON9 = "o" if self.S > 3 else "x"
        SCON10 = "o" if self.popular < 3 or self.popular > 8 else "x"
        SCON11 = "o" if self.H < 2 else "x"
        SCON12 = "o" if self.R < 17 else "x"
        self.SCON = SCON1 + SCON2 + SCON3 + SCON4 + SCON5 + SCON6 + SCON7 + SCON8 + SCON9 + SCON10 + SCON11 + SCON12
        self.sumSCON = self.SCON.count("o")  # 8이상, 14세 이상만 고려
        self.HVI_premise = self.sum_T == 0
        HVI2 = "o" if self.Zf > 12 else "x"
        HVI3 = "o" if self.Zd > 3.5 else "x"
        HVI4 = "o" if self.S > 3 else "x"
        HVI5 = "o" if self.human_cont > 6 else "x"
        HVI6 = "o" if self.H_paren + self.A_paren + self.Hd_paren + self.Ad_paren > 3 else "x"
        try:
            ratio = (self.H + self.A) / (self.Hd + self.Ad)
            HVI7 = "o" if ratio < 4 else "x"
        except ZeroDivisionError:
            # 0으로 나누는 에러가 발생한 경우
            HVI7 = "x"
            self.HVI_except = "{}:{}".format((self.H + self.A), (self.Hd + self.Ad))
        HVI8 = "o" if self.Cg > 3 else "x"
        self.HVI = HVI2 + HVI3 + HVI4 + HVI5 + HVI6 + HVI7 + HVI8
        self.sumHVI = self.HVI.count('o')  # 4이상
        OBS1 = "o" if self.Dd > 3 else "x"
        OBS2 = "o" if self.Zf > 12 else "x"
        OBS3 = "o" if self.Zd > 3.0 else "x"
        OBS4 = "o" if self.popular > 7 else "x"
        OBS5 = "o" if self.fqx_plus > 1 else "x"
        OBS6 = "o" if OBS1 == OBS2 == OBS3 == OBS4 == OBS5 == "o" else "x"
        OBS7 = "o" if sum(1 for obs in [OBS1, OBS2, OBS3, OBS4] if obs == "o") >= 2 and self.fqx_plus > 3 else "x"
        OBS8 = "o" if (sum(1 for obs in [OBS1, OBS2, OBS3, OBS4, OBS5] if obs == "o") >= 3 and
                       self.x_plus_per > 0.89) else "x"
        OBS9 = "o" if self.fqx_plus > 3 and self.x_plus_per > 0.89 else "x"
        self.OBS = OBS1 + OBS2 + OBS3 + OBS4 + OBS5 + OBS6 + OBS7 + OBS8 + OBS9
        self.OBS_posi = "o" in self.OBS[-4:]

    def save(self, *args, **kwargs):
        if not self.pk:
            self.calculate_values()
        super(StructuralSummary, self).save(*args, **kwargs)

    def __str__(self):
        return str(self.client)
