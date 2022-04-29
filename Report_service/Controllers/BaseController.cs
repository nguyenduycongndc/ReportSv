using Report_service.Attributes;
using Report_service.DataAccess;
using Report_service.Repositories;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Report_service.Controllers
{

    [BaseAuthorize]
    public class BaseController : ControllerBase
    {
        protected readonly ILoggerManager _logger;
        protected readonly IUnitOfWork _uow;
        //protected readonly IMapper _mapper;

        public BaseController(ILoggerManager logger,IUnitOfWork uow) : base()
        {
            _uow = uow;
            _logger = logger;
        }
    }
}
