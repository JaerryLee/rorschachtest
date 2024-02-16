# views.py
from .models import Post, Comment, Notice
from functools import wraps
from django.http import HttpResponseForbidden
from django.shortcuts import render, get_object_or_404, redirect
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.contrib.auth.decorators import login_required
from .forms import PostForm, CommentForm


def group_required(group_name):
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return login_required(view_func)(request, *args, **kwargs)

            if request.user.group == group_name:
                return view_func(request, *args, **kwargs)
            else:
                return HttpResponseForbidden("중급 이상 이수자만 접속 가능한 페이지입니다.")

        return _wrapped_view

    return decorator


@login_required
def beginner_board(request):
    # Retrieve all posts
    posts = Post.objects.filter(group='beginner').order_by('-created_at')

    # Search functionality
    search_query = request.GET.get('search', '')
    if search_query:
        posts = posts.filter(title__icontains=search_query)

    # Pagination
    paginator = Paginator(posts, 10)  # Show 10 posts per page
    page = request.GET.get('page')
    try:
        posts = paginator.page(page)
    except PageNotAnInteger:
        posts = paginator.page(1)
    except EmptyPage:
        posts = paginator.page(paginator.num_pages)

    return render(request, 'beginner_board.html', {'posts': posts, 'search_query': search_query})


@group_required('intermediate')
def intermediate_board(request):
    # Retrieve all posts
    posts = Post.objects.filter(group='intermediate').order_by('-created_at')

    # Search functionality
    search_query = request.GET.get('search', '')
    if search_query:
        posts = posts.filter(title__icontains=search_query)

    # Pagination
    paginator = Paginator(posts, 10)  # Show 10 posts per page
    page = request.GET.get('page')
    try:
        posts = paginator.page(page)
    except PageNotAnInteger:
        posts = paginator.page(1)
    except EmptyPage:
        posts = paginator.page(paginator.num_pages)

    return render(request, 'intermediate_board.html', {'posts': posts, 'search_query': search_query})


@login_required
def post_detail(request, post_id):
    post = get_object_or_404(Post, id=post_id)
    if request.user.group != 'intermediate' and post.group == 'intermediate':
        return HttpResponseForbidden("중급 이상 이수자만 접속 가능한 페이지입니다.")
    comments = Comment.objects.filter(post=post)

    can_delete = False
    if request.user == post.author:
        can_delete = True

    if request.method == 'POST':
        form = CommentForm(request.POST)
        if form.is_valid():
            comment = form.save(commit=False)
            comment.post = post
            comment.author = request.user  # Assuming you have user authentication
            comment.save()
            return redirect('board:post_detail', post_id=post_id)
    else:
        form = CommentForm()

    return render(request, 'post_detail.html', {'post': post, 'comments': comments, 'form': form,
                                                'can_delete': can_delete})


@login_required
def create_post(request, group):
    if request.method == 'POST':
        form = PostForm(request.POST)
        if form.is_valid():
            new_post = form.save(commit=False)
            new_post.author = request.user
            new_post.group = group
            new_post.save()
            return redirect('board:beginner_board' if group == 'beginner' else 'board:intermediate_board')
    else:
        form = PostForm()

    return render(request, 'create_post.html', {'form': form, 'group': group})


@login_required
def delete_post(request, post_id):
    post = get_object_or_404(Post, id=post_id)
    group = post.group

    if request.user == post.author:
        post.delete()

    return redirect('board:beginner_board' if group == 'beginner' else 'board:intermediate_board')


def notice(request):
    # Retrieve all notices
    notices = Notice.objects.all()

    # Search functionality
    search_query = request.GET.get('search', '')
    if search_query:
        notices = notices.filter(title__icontains=search_query)

    # Pagination
    paginator = Paginator(notices, 10)  # Show 10 posts per page
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
