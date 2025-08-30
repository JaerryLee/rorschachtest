from functools import wraps

from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.http import HttpResponseBadRequest, HttpResponseForbidden
from django.shortcuts import render, get_object_or_404, redirect

from .forms import PostForm, CommentForm
from .models import Post, Comment, Notice

GROUP_LEVEL = {'beginner': 1, 'intermediate': 2, 'advanced': 3}
GROUP_LABEL = {'beginner': '초급', 'intermediate': '중급', 'advanced': '고급'}


def has_min_group(user, min_group: str) -> bool:
    """user가 min_group 이상 권한인지 확인"""
    if not user.is_authenticated:
        return False
    return GROUP_LEVEL.get(getattr(user, 'group', None), 0) >= GROUP_LEVEL[min_group]


def group_min_required(min_group_name: str):
    """해당 등급 이상만 접근 가능 데코레이터"""
    def decorator(view_func):
        @wraps(view_func)
        @login_required
        def _wrapped_view(request, *args, **kwargs):
            if has_min_group(request.user, min_group_name):
                return view_func(request, *args, **kwargs)
            return HttpResponseForbidden(
                f"{GROUP_LABEL[min_group_name]} 이상 이수자만 접속 가능한 페이지입니다."
            )
        return _wrapped_view
    return decorator


@login_required
def beginner_board(request):
    posts = Post.objects.filter(group='beginner').order_by('-created_at')

    search_query = request.GET.get('search', '')
    if search_query:
        posts = posts.filter(title__icontains=search_query)

    paginator = Paginator(posts, 10)
    page = request.GET.get('page')
    try:
        posts = paginator.page(page)
    except PageNotAnInteger:
        posts = paginator.page(1)
    except EmptyPage:
        posts = paginator.page(paginator.num_pages)

    return render(request, 'beginner_board.html', {'posts': posts, 'search_query': search_query})


@group_min_required('intermediate')
def intermediate_board(request):
    posts = Post.objects.filter(group='intermediate').order_by('-created_at')

    search_query = request.GET.get('search', '')
    if search_query:
        posts = posts.filter(title__icontains=search_query)

    paginator = Paginator(posts, 10)
    page = request.GET.get('page')
    try:
        posts = paginator.page(page)
    except PageNotAnInteger:
        posts = paginator.page(1)
    except EmptyPage:
        posts = paginator.page(paginator.num_pages)

    return render(request, 'intermediate_board.html', {'posts': posts, 'search_query': search_query})


@group_min_required('advanced')
def advanced_board(request):
    posts = Post.objects.filter(group='advanced').order_by('-created_at')

    search_query = request.GET.get('search', '')
    if search_query:
        posts = posts.filter(title__icontains=search_query)

    paginator = Paginator(posts, 10)
    page = request.GET.get('page')
    try:
        posts = paginator.page(page)
    except PageNotAnInteger:
        posts = paginator.page(1)
    except EmptyPage:
        posts = paginator.page(paginator.num_pages)

    return render(request, 'advanced_board.html', {'posts': posts, 'search_query': search_query})


@login_required
def post_detail(request, post_id):
    post = get_object_or_404(Post, id=post_id)

    # 게시글 그룹 접근 권한 재검증 (URL 직행 방지)
    if post.group in GROUP_LEVEL:
        user_level = GROUP_LEVEL.get(getattr(request.user, 'group', None), 0)
        required_level = GROUP_LEVEL[post.group]
        if user_level < required_level:
            return HttpResponseForbidden(
                f"{GROUP_LABEL[post.group]} 이상 이수자만 접속 가능한 페이지입니다."
            )

    comments = Comment.objects.filter(post=post)

    can_delete = request.user == post.author

    if request.method == 'POST':
        form = CommentForm(request.POST)
        if form.is_valid():
            comment = form.save(commit=False)
            comment.post = post
            comment.author = request.user
            comment.save()
            return redirect('board:post_detail', post_id=post_id)
    else:
        form = CommentForm()

    return render(
        request,
        'post_detail.html',
        {'post': post, 'comments': comments, 'form': form, 'can_delete': can_delete},
    )


@login_required
def create_post(request, group):
    # 1) 유효한 그룹인지 확인
    if group not in GROUP_LEVEL:
        return HttpResponseBadRequest("잘못된 게시판입니다.")

    # 2) 현재 사용자 권한 검증 (중급/고급 글쓰기 차단)
    if not has_min_group(request.user, group):
        return HttpResponseForbidden(
            f"{GROUP_LABEL[group]} 이상 이수자만 글을 작성할 수 있습니다."
        )

    if request.method == 'POST':
        form = PostForm(request.POST, request.FILES)
        if form.is_valid():
            new_post = form.save(commit=False)
            new_post.author = request.user
            # 폼에서 group을 받지 않고 서버에서 강제 설정(우회 방지)
            new_post.group = group
            new_post.save()
            return redirect(
                'board:beginner_board' if group == 'beginner'
                else 'board:intermediate_board' if group == 'intermediate'
                else 'board:advanced_board'
            )
    else:
        form = PostForm()

    return render(request, 'create_post.html', {'form': form, 'group': group})


@login_required
def delete_post(request, post_id):
    post = get_object_or_404(Post, id=post_id)
    group = post.group

    if request.user == post.author:
        post.delete()

    return redirect(
        'board:beginner_board' if group == 'beginner'
        else 'board:intermediate_board' if group == 'intermediate'
        else 'board:advanced_board'
    )


def notice(request):
    notices = Notice.objects.all().order_by('-created_at')

    search_query = request.GET.get('search', '')
    if search_query:
        notices = notices.filter(title__icontains=search_query)

    paginator = Paginator(notices, 10)
    page = request.GET.get('page')
    try:
        notices = paginator.page(page)
    except PageNotAnInteger:
        notices = paginator.page(1)
    except EmptyPage:
        notices = paginator.page(paginator.num_pages)

    return render(request, 'notice.html', {'notices': notices, 'search_query': search_query})


def notice_detail(request, notice_id):
    notice = get_object_or_404(Notice, id=notice_id)
    return render(request, 'notice_detail.html', {'notice': notice})
