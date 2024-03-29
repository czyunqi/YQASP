    <nav class="navbar navbar-default navbar-fixed-top" role="navigation">
      <div class="container">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#navbar-collapse-1">
            <span class="sr-only">切换导航</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
          <a class="navbar-brand" href="#">YunqASP <small>v3</small></a> 
        </div>

        <div class="collapse navbar-collapse" id="navbar-collapse-1">
          <ul class="nav navbar-nav">
            <li><a href="http://www.yqasp.cn/overview"><span class="glyphicon glyphicon-eye-open"></span> 概览</a></li>
            <li><a href="http://www.yqasp.cn/tutorials"><span class="glyphicon glyphicon-tint"></span> 如何使用</a></li>
            <li><a href="http://www.yqasp.cn/api"><span class="glyphicon glyphicon-book"></span> API文档</a></li>
            <li class="dropdown active">
              <a href="#" class="dropdown-toggle" data-toggle="dropdown"><span class="glyphicon glyphicon-briefcase"></span> 示例 <b class="caret"></b></a>
              <ul class="dropdown-menu">
                <li class="dropdown-header">YunqASP核心类</li>
                <li><a href="http://www.yqasp.cn/api">字符串(YQasp.Str)</a></li>
                <li><a href="http://www.yqasp.cn/api">数据库(YQasp.Db)</a></li>
                <li><a href="http://www.yqasp.cn/api">模板引擎(YQasp.Tpl)</a></li>
                <li><a href="http://www.yqasp.cn/api">JSON数据(YQasp.Json)</a></li>
                <li class="divider"></li>
                <li class="dropdown-header">YunqASP插件</li>
                <li><a href="http://www.yqasp.cn/plugins">MD5加密</a></li>
              </ul>
            </li><!-- /dropdown -->
          </ul><!-- /.navbar-nav -->
          <a href="http://www.yqasp.cn/donate" role="button" class="btn btn-success navbar-btn navbar-right m-l-10" target="_blank"><span class="glyphicon glyphicon-heart"></span> 捐赠</a>
          <form class="navbar-form navbar-right" role="search" action="http://www.yqasp.cn/start" method="get" target="_blank">
            <div class="form-group">
              <input type="text" name="id" class="form-control" placeholder="搜索...">
              <input type="hidden" name="do" value="search">
            </div>
            <button type="submit" class="btn btn-default">
              <span class="glyphicon glyphicon-search"></span>
            </button>
          </form>
        </div><!-- /.navbar-collapse -->
      </div><!-- /.container -->
    </nav>